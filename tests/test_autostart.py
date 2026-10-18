"""autostart.py's update: main pulled only when it's safe (on main, no changes of its own, a fast-forward), the
requirements installed only when they changed, and the app started either way. Throwaway git repos only."""
import os
import subprocess
import sys
from types import SimpleNamespace

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import autostart  # noqa: E402


def git(repo, *args):
    return subprocess.run(["git", "-C", str(repo), "-c", "user.name=Test", "-c", "user.email=t@example.com", *args],
                          capture_output=True, text=True, check=True).stdout.strip()


class Log:
    def __init__(self):
        self.lines = []

    def write(self, line):
        self.lines.append(line)


@pytest.fixture()
def repos(tmp_path):
    """origin (GitHub), dev (someone pushing to main) and here (the copy auto-start runs)."""
    origin, dev, here = tmp_path / "origin.git", tmp_path / "dev", tmp_path / "here"
    subprocess.run(["git", "init", "--bare", "-b", "main", str(origin)], check=True, capture_output=True)
    subprocess.run(["git", "clone", str(origin), str(dev)], check=True, capture_output=True)
    git(dev, "checkout", "-b", "main")
    (dev / "app.py").write_text("PORT = 5001\n")
    (dev / "requirements.txt").write_text("Flask\n")
    git(dev, "add", ".")
    git(dev, "commit", "-m", "first")
    git(dev, "push", "origin", "main")
    subprocess.run(["git", "clone", str(origin), str(here)], check=True, capture_output=True)
    return origin, dev, here


def push(dev, name, text, message):
    (dev / name).write_text(text)
    git(dev, "add", ".")
    git(dev, "commit", "-m", message)
    git(dev, "push", "origin", "main")


def run_update(here, tmp_path, ok=True):
    calls = []

    def pip():
        calls.append(1)
        return SimpleNamespace(returncode=0 if ok else 1, stdout="", stderr="ERROR: no internet")
    log = Log()
    result = autostart.update(log, repo=here, requirements=here / "requirements.txt", stamp=tmp_path / "stamp",
                              run_pip=pip)
    return result, calls, log


def test_main_is_pulled_and_the_requirements_installed_only_when_they_change(repos, tmp_path):
    _, dev, here = repos
    push(dev, "app.py", "PORT = 5001\nprint('new')\n", "a fix")
    result, pip, log = run_update(here, tmp_path)
    assert result == "updated" and "print('new')" in (here / "app.py").read_text()
    assert pip == [1] and any("a fix" in x for x in log.lines)  # first start: installed once
    assert run_update(here, tmp_path)[:2] == ("up to date", [])  # nothing new, requirements the same: no pip
    push(dev, "requirements.txt", "Flask\nrequests\n", "needs requests")
    result, pip, _ = run_update(here, tmp_path)
    assert result == "updated" and pip == [1]


def test_a_copy_with_its_own_changes_or_on_another_branch_is_left_as_it_is(repos, tmp_path):
    _, dev, here = repos
    push(dev, "app.py", "PORT = 5001\nprint('new')\n", "a fix")
    (here / "app.py").write_text("PORT = 5001\n# work in progress\n")
    result, _, _ = run_update(here, tmp_path)
    assert result.startswith("skipped: this copy has changes of its own")
    assert "work in progress" in (here / "app.py").read_text()  # untouched
    git(here, "checkout", "--", "app.py")
    git(here, "checkout", "-b", "feature")
    assert run_update(here, tmp_path)[0].startswith("skipped: this copy is on branch 'feature'")


def test_main_that_went_another_way_here_is_never_overwritten(repos, tmp_path):
    _, dev, here = repos
    push(dev, "app.py", "PORT = 5001\nprint('origin')\n", "on GitHub")
    (here / "local.txt").write_text("mine")
    git(here, "add", ".")
    git(here, "commit", "-m", "a commit only here")
    result, _, _ = run_update(here, tmp_path)
    assert result.startswith("skipped: main here and on GitHub have gone different ways")
    assert (here / "local.txt").exists() and "a commit only here" in git(here, "log", "-1", "--format=%s")


def test_offline_or_pip_failing_still_starts_with_what_it_has(repos, tmp_path):
    origin, dev, here = repos
    git(here, "remote", "set-url", "origin", str(tmp_path / "nowhere.git"))
    assert run_update(here, tmp_path)[0].startswith("skipped: couldn't reach GitHub")
    git(here, "remote", "set-url", "origin", str(origin))
    result, pip, log = run_update(here, tmp_path, ok=False)
    assert result == "up to date" and pip == [1] and any("pip failed" in x for x in log.lines)
    assert not (tmp_path / "stamp").exists()  # tried again at the next start
