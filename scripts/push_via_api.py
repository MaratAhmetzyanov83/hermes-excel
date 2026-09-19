#!/usr/bin/env python3
"""Заливает текущее содержимое репозитория в GitHub через API (вместо `git push`).

Зачем: на некоторых машинах (VPN/DPI) TLS у `git push` рвётся — `schannel: failed to receive
handshake` или `unexpected eof while reading`, — а `gh`/API по тому же адресу работает.
Скрипт собирает коммит средствами Git Data API и двигает ветку.

Использование:
  python scripts/push_via_api.py --repo <owner>/<name> [--branch main] [--message "..."] [--dry-run]
"""
from __future__ import annotations

import argparse
import base64
import json
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BINARY_EXT = {".png", ".jpg", ".jpeg", ".gif", ".ico", ".xlsx", ".zip", ".pdf"}


def gh_api(path: str, method: str = "GET", body: dict | None = None, allow_fail: bool = False,
           attempts: int = 4):
    """gh api с повторами: на VPN-каналах запросы к api.github.com периодически рвутся (EOF)."""
    last = ""
    for n in range(1, attempts + 1):
        cmd = ["gh", "api", "-X", method, path]
        if body is not None:
            tmp = ROOT / ".git" / "api-body.json"
            tmp.parent.mkdir(exist_ok=True)
            tmp.write_text(json.dumps(body), encoding="utf-8")
            cmd += ["--input", str(tmp)]
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if proc.returncode == 0:
            return json.loads(proc.stdout) if proc.stdout.strip() else None
        last = (proc.stderr or proc.stdout).strip().splitlines()[-1] if (proc.stderr or proc.stdout) else ""
        if n < attempts:
            print(f"    попытка {n} не прошла ({last[:60]}) — повтор через {2 * n} с")
            time.sleep(2 * n)
    if allow_fail:
        return None
    sys.exit(f"gh api {method} {path} → не удалось за {attempts} попыток: {last}")


def file_bytes(rel: str) -> bytes:
    """Байты ровно такие, какие отправил бы `git push`: содержимое из индекса
    (с применёнными фильтрами .gitattributes), а не то, что лежит на диске."""
    proc = subprocess.run(["git", "cat-file", "blob", f":{rel}"], cwd=ROOT, capture_output=True)
    if proc.returncode == 0 and proc.stdout:
        return proc.stdout
    return (ROOT / rel).read_bytes()


def tracked_files() -> list[str]:
    out = subprocess.run(["git", "ls-files"], cwd=ROOT, capture_output=True, text=True, check=True)
    return [line for line in out.stdout.splitlines() if line.strip()]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--repo", required=True, help="owner/name")
    ap.add_argument("--branch", default="main")
    ap.add_argument("--message", default=None)
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    files = tracked_files()
    if not files:
        sys.exit("в индексе git нет файлов — сначала `git add -A`")

    message = a.message
    if not message:
        message = subprocess.run(["git", "log", "-1", "--pretty=%B"], cwd=ROOT,
                                 capture_output=True, text=True).stdout.strip() or "initial commit"

    if a.dry_run:
        print(f"будет залито {len(files)} файлов в {a.repo}#{a.branch}:\n  " + "\n  ".join(files))
        return

    print(f"1/4 ветка…")
    head = gh_api(f"repos/{a.repo}/git/ref/heads/{a.branch}", allow_fail=True)
    if head is None:
        # В пустом репозитории Git Data API отдаёт 409 «Git Repository is empty»,
        # поэтому заводим первый коммит через Contents API — он и создаёт ветку.
        boot = "README.md" if "README.md" in files else files[0]
        first = boot.replace("\\", "/")
        data = file_bytes(boot)
        gh_api(f"repos/{a.repo}/contents/{first}", "PUT",
               {"message": message, "content": base64.b64encode(data).decode("ascii")})
        head = gh_api(f"repos/{a.repo}/git/ref/heads/{a.branch}")
        print("    заведён первый коммит:", head["object"]["sha"][:8])
    else:
        print("    ветка уже есть:", head["object"]["sha"][:8])

    print(f"2/4 блобы ({len(files)} шт.)…")
    tree = []
    for rel in files:
        data = file_bytes(rel)
        payload = {"content": base64.b64encode(data).decode("ascii"), "encoding": "base64"}
        blob = gh_api(f"repos/{a.repo}/git/blobs", "POST", payload)
        tree.append({"path": rel.replace("\\", "/"), "mode": "100644", "type": "blob", "sha": blob["sha"]})
        print(f"    {rel} → {blob['sha'][:8]}")

    print("3/4 дерево и коммит…")
    tree_obj = gh_api(f"repos/{a.repo}/git/trees", "POST", {"tree": tree})
    commit_body = {"message": message, "tree": tree_obj["sha"],
                   "parents": [head["object"]["sha"]] if head else []}
    commit = gh_api(f"repos/{a.repo}/git/commits", "POST", commit_body)
    print("    tree:", tree_obj["sha"][:8], "| commit:", commit["sha"][:8])

    print("4/4 двигаю ветку…")
    gh_api(f"repos/{a.repo}/git/refs/heads/{a.branch}", "PATCH", {"sha": commit["sha"]})
    print(f"\nготово: https://github.com/{a.repo}/tree/{a.branch}")


if __name__ == "__main__":
    main()
