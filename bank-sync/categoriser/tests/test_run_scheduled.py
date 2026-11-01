"""scripts/run-scheduled.sh auto-update, against a local git repo and a fake `docker` (no network)."""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

import pytest

CATEGORISER = Path(__file__).resolve().parents[1]
SUB = "bank-sync/categoriser"

FAKE_DOCKER = r"""#!/bin/bash
# Image labels live in $STATE/labels (tree, commit); every build/run is appended to $STATE/calls.
S=$STATE
case "$1" in
  info) exit 0 ;;
  image)
    [ -f "$S/labels" ] || exit 1
    if [ "$3" = -f ]; then case "$4" in *tree*) sed -n 1p "$S/labels" ;; *) sed -n 2p "$S/labels" ;; esac; fi ;;
  build)
    shift; tree=""; commit=""
    while [ $# -gt 0 ]; do
      case "$1" in
        --label) case "$2" in *tree=*) tree=${2#*=} ;; *commit=*) commit=${2#*=} ;; esac; shift ;;
        -*) ;;
        *) ctx=$1 ;;
      esac; shift
    done
    [ -z "$FAIL_BUILD" ] && [ -f "$ctx/Dockerfile" ] || exit 1
    printf '%s\n%s\n' "$tree" "$commit" > "$S/labels"; echo build >> "$S/calls" ;;
  compose) echo run >> "$S/calls" ;;
esac
"""


def git(cwd: Path, *args: str) -> str:
    return subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@example.com", *args], cwd=cwd,
                          check=True, capture_output=True, text=True).stdout.strip()


@pytest.fixture
def env(tmp_path):
    if not shutil.which("git"):
        pytest.skip("git not installed")
    upstream = tmp_path / "upstream"
    target = upstream / SUB
    shutil.copytree(CATEGORISER, target, ignore=shutil.ignore_patterns(".venv", "__pycache__", "*.egg-info"))
    git(upstream, "init", "-q", "-b", "main")
    git(upstream, "add", "-A")
    git(upstream, "commit", "-qm", "init")

    bin_dir, home, runtime = tmp_path / "bin", tmp_path / "home", tmp_path / "runtime"
    for d in (bin_dir, home / "Library/LaunchAgents", runtime):
        d.mkdir(parents=True)
    (bin_dir / "docker").write_text(FAKE_DOCKER)
    (bin_dir / "docker").chmod(0o755)
    shutil.copy2(CATEGORISER / "scripts/run-scheduled.sh", runtime / "run-scheduled.sh")
    (runtime / ".env").write_text("SPREADSHEET_ID=x\n")
    plist = (CATEGORISER / "launchd/lu.niklas.categoriser.plist.template").read_text()
    (home / "Library/LaunchAgents/lu.niklas.categoriser.plist").write_text(
        plist.replace("__CATEGORISER_DIR__", str(runtime)))
    environ = {**os.environ, "PATH": f"{bin_dir}:{os.environ['PATH']}", "HOME": str(home), "STATE": str(tmp_path),
               "CATEGORISER_REPO_URL": f"file://{upstream}"}

    class Env:
        def run(self, **extra) -> tuple[str, list[str]]:
            (runtime / "logs/.last-success").unlink(missing_ok=True)
            (tmp_path / "calls").unlink(missing_ok=True)
            out = subprocess.run(["bash", str(runtime / "run-scheduled.sh")], env={**environ, **extra},
                                 check=True, capture_output=True, text=True).stdout
            calls = tmp_path / "calls"
            return out, calls.read_text().split() if calls.exists() else []

        def commit(self, path: str, text: str) -> str:
            with open(upstream / path, "a") as f:
                f.write(text)
            git(upstream, "commit", "-qam", "change")
            return git(upstream, "rev-parse", "--short=7", "HEAD")

        upstream_dir = upstream
        runtime_dir = runtime

    return Env()


def test_builds_when_categoriser_changed_and_not_otherwise(env):
    out, calls = env.run()
    assert calls == ["build", "run"] and "image updated to" in out
    assert env.run()[1] == ["run"]                                    # unchanged: no rebuild

    (env.upstream_dir / "unrelated.txt").write_text("x")
    git(env.upstream_dir, "add", "-A")
    git(env.upstream_dir, "commit", "-qm", "outside")
    assert env.run()[1] == ["run"]                                    # change outside bank-sync/categoriser

    commit = env.commit(f"{SUB}/README.md", "\nmore\n")
    out, calls = env.run()
    assert calls == ["build", "run"] and f"run start ({commit})" in out


def test_failed_build_or_fetch_still_runs_the_current_image(env):
    env.run()
    env.commit(f"{SUB}/README.md", "\nmore\n")
    out, calls = env.run(FAIL_BUILD="1")
    assert calls == ["run"] and "build failed: keeping the current image" in out
    shutil.rmtree(env.upstream_dir / ".git")
    out, calls = env.run()
    assert calls == ["run"] and "auto-update did not complete" in out


def test_refreshes_itself_and_warns_about_a_changed_schedule(env):
    env.run()
    env.commit(f"{SUB}/scripts/run-scheduled.sh", "\n# newer\n")
    env.commit(f"{SUB}/launchd/lu.niklas.categoriser.plist.template", "\n<!-- newer -->\n")
    out, _ = env.run()
    assert "run-scheduled.sh updated" in out and "WARNING: the launchd schedule changed" in out
    assert (env.runtime_dir / "run-scheduled.sh").read_text().endswith("# newer\n")
    assert os.access(env.runtime_dir / "run-scheduled.sh", os.X_OK)


def test_opt_out_keeps_the_installed_image(env):
    env.run()
    env.commit(f"{SUB}/README.md", "\nmore\n")
    (env.runtime_dir / ".env").write_text("SPREADSHEET_ID=x\nCATEGORISER_AUTO_UPDATE=0\n")
    out, calls = env.run()
    assert "auto-update off" in out and calls == ["run"]
