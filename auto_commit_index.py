import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parent
INDEX_FILE = ROOT / "index.html"
POLL_SECONDS = 2.0
DEBOUNCE_SECONDS = 4.0


def run_git(args):
    return subprocess.run(
        ["git", *args],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
    )


def has_index_changes():
    p = run_git(["status", "--porcelain", "--", "index.html"])
    return bool((p.stdout or "").strip())


def commit_and_push():
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"Auto-update index.html ({ts})"

    add = run_git(["add", "--", "index.html"])
    if add.returncode != 0:
        return False, (add.stderr or add.stdout or "git add falhou").strip()

    commit = run_git(["commit", "-m", msg, "--", "index.html"])
    out = ((commit.stdout or "") + "\n" + (commit.stderr or "")).strip()
    if commit.returncode != 0:
        if "nothing to commit" in out.lower():
            return True, "Sem alteracoes para commit."
        return False, out or "git commit falhou"

    push = run_git(["push"])
    push_out = ((push.stdout or "") + "\n" + (push.stderr or "")).strip()
    if push.returncode != 0:
        return False, push_out or "git push falhou"
    return True, "Commit + push concluido."


def main():
    if not INDEX_FILE.exists():
        print("index.html nao encontrado no diretorio do projeto.")
        return 1

    print("Auto-commit de index.html ativo.")
    print(f"Monitorando: {INDEX_FILE}")

    last_mtime = INDEX_FILE.stat().st_mtime
    pending_since = None

    while True:
        try:
            now_mtime = INDEX_FILE.stat().st_mtime
        except FileNotFoundError:
            time.sleep(POLL_SECONDS)
            continue

        if now_mtime != last_mtime:
            last_mtime = now_mtime
            pending_since = time.time()

        if pending_since is not None:
            if (time.time() - pending_since) >= DEBOUNCE_SECONDS:
                if has_index_changes():
                    ok, info = commit_and_push()
                    stamp = datetime.now().strftime("%H:%M:%S")
                    status = "OK" if ok else "ERRO"
                    print(f"[{stamp}] {status}: {info}")
                pending_since = None

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("Encerrado.")
        raise SystemExit(0)
