"""
Auto-start for the Warehouse app: a Windows scheduled task that, every time you log in to Windows, brings the app up to
date from GitHub's main branch, installs what it needs, and starts it in the background: no console window to keep
open, nothing to remember after a restart.

Manage it with the venv's python, from this folder:

    venv\\Scripts\\python autostart.py install     set up the task, update and start the app
    venv\\Scripts\\python autostart.py status      is it running? is auto-start on? what version?
    venv\\Scripts\\python autostart.py update      just update (git pull of main + requirements), don't restart
    venv\\Scripts\\python autostart.py restart     stop the app, update, start it again
    venv\\Scripts\\python autostart.py stop        stop the app and pause auto-start
    venv\\Scripts\\python autostart.py uninstall   stop the app and remove the task

The app's output goes to logs\\server.log (the app's own log, logs\\warehouse.log, is on its Logs page as before).

How it works: the task runs `pythonw autostart.py run` at logon and again every minute after that. `run` exits straight
away if something is already answering on the app's port, so a running app is left alone (one started by hand with
Start_WarehouseApp.bat too) and one that stopped is back within a minute. Otherwise it updates, then starts the same
`venv\\Scripts\\python.exe app.py` the .bat does, without a window.

The update is careful, because this folder may be where someone is working on the code:
  - only when this copy is on the main branch with no changes of its own (git status clean, untracked files aside):
    anything else is left exactly as it is, and the log says why it wasn't updated;
  - only a fast-forward (git merge --ff-only): never a reset, a forced checkout or a merge commit;
  - the requirements are installed again only when requirements.txt changed (pip), so a normal start stays quick;
  - offline, or anything failing: logged, and the app starts with what it has.
Anyone who can push to main can change what this PC runs at its next start: keep main for code that's ready.

It starts at logon, not at boot: a task that runs before anyone logs in needs admin rights to create, and Excel (which
Packing Sheets drives) needs a logged-in desktop anyway.
"""
import csv
import hashlib
import os
import re
import socket
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

BASE_DIR = Path(__file__).resolve().parent
TASK_NAME = "Warehouse App"
BRANCH = "main"
RECHECK_MINUTES = 1
VENV_PYTHON = BASE_DIR / "venv" / "Scripts" / "python.exe"
VENV_PYTHONW = BASE_DIR / "venv" / "Scripts" / "pythonw.exe"
REQUIREMENTS = BASE_DIR / "requirements.txt"
REQUIREMENTS_STAMP = BASE_DIR / "venv" / ".requirements-installed"  # the requirements.txt pip last installed
LOG_DIR = BASE_DIR / "logs"
LOG_PATH = LOG_DIR / "server.log"
MAX_LOG_BYTES = 5 * 1024 * 1024  # an app left running for weeks keeps at most about twice this
NO_WINDOW = getattr(subprocess, 'CREATE_NO_WINDOW', 0)


def _app_port():
    """The port app.py serves on (its PORT = …), so the two never disagree."""
    m = re.search(r'^PORT\s*=\s*(\d+)', (BASE_DIR / "app.py").read_text(encoding="utf-8"), re.M)
    return int(m.group(1)) if m else 5001


PORT = _app_port()


def _is_serving():
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=2):
            return True
    except OSError:
        return False


class _Log:
    """logs/server.log, one timestamped line at a time; past MAX_LOG_BYTES it moves to server.log.1 (replacing the one
    before) and a fresh file starts."""

    def __init__(self, path=LOG_PATH, quiet=False):
        self.path = Path(path)
        self.path.parent.mkdir(exist_ok=True)
        self._file = open(self.path, "a", encoding="utf-8")
        self.quiet = quiet

    def write(self, line):
        self._file.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {line.rstrip()}\n")
        self._file.flush()
        if not self.quiet and sys.stdout:
            print(line.rstrip())
        if self._file.tell() > MAX_LOG_BYTES:
            self._file.close()
            try:
                os.replace(self.path, self.path.with_suffix(".log.1"))
            except OSError:
                pass  # locked by something reading it: keep appending
            self._file = open(self.path, "a", encoding="utf-8")


def _git(repo, *args, timeout=60):
    return subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True, timeout=timeout,
                          stdin=subprocess.DEVNULL, creationflags=NO_WINDOW)


def _requirements_hash(path):
    try:
        return hashlib.sha1(Path(path).read_bytes()).hexdigest()
    except OSError:
        return ''


def update(log, repo=BASE_DIR, python=VENV_PYTHON, requirements=REQUIREMENTS, stamp=REQUIREMENTS_STAMP, run_pip=None):
    """Brings this copy up to date with origin/main when that's safe, then the requirements when they changed.
    Returns what happened: 'updated', 'up to date', or 'skipped: <why>'. Never raises: the app starts either way."""
    try:
        if _git(repo, "--version").returncode != 0:
            return _skip(log, "git isn't installed or not on the PATH")
    except (OSError, subprocess.SubprocessError):
        return _skip(log, "git isn't installed or not on the PATH")
    try:
        branch = _git(repo, "rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
        if branch != BRANCH:
            return _skip(log, f"this copy is on branch '{branch}', not {BRANCH} (someone is working on it): left as it is")
        changed = _git(repo, "status", "--porcelain", "--untracked-files=no").stdout.strip()
        if changed:
            return _skip(log, f"this copy has changes of its own ({len(changed.splitlines())} file(s)): left as it is")
        before = _git(repo, "rev-parse", "--short", "HEAD").stdout.strip()
        fetched = _git(repo, "fetch", "--quiet", "origin", BRANCH, timeout=120)
        if fetched.returncode != 0:
            return _skip(log, f"couldn't reach GitHub ({(fetched.stderr or fetched.stdout).strip()[:200]}): starting as it is")
        merged = _git(repo, "merge", "--ff-only", "--quiet", f"origin/{BRANCH}")
        if merged.returncode != 0:
            return _skip(log, f"main here and on GitHub have gone different ways ({(merged.stderr or merged.stdout).strip()[:200]}): "
                              "left as it is")
        after = _git(repo, "rev-parse", "--short", "HEAD").stdout.strip()
        result = 'up to date' if before == after else 'updated'
        if result == 'updated':
            subjects = _git(repo, "log", "--format=%h %s", f"{before}..{after}").stdout.strip().splitlines()
            log.write(f"[update] {before} -> {after}: {len(subjects)} change(s)")
            for s in subjects[:20]:
                log.write(f"[update]   {s}")
        else:
            log.write(f"[update] up to date with GitHub's {BRANCH} ({after})")
    except (OSError, subprocess.SubprocessError) as e:
        return _skip(log, f"git failed ({e}): starting as it is")
    _requirements(log, python, requirements, stamp, run_pip)
    return result


def _skip(log, why):
    log.write(f"[update] not updated: {why}")
    return f"skipped: {why}"


def _requirements(log, python, requirements, stamp, run_pip=None):
    """pip install -r requirements.txt, only when it changed since the last install that worked."""
    wanted = _requirements_hash(requirements)
    if not wanted:
        return
    try:
        if Path(stamp).read_text(encoding="utf-8").strip() == wanted:
            return
    except OSError:
        pass
    log.write("[update] requirements.txt changed: installing them (pip)")
    run_pip = run_pip or (lambda: subprocess.run([str(python), "-m", "pip", "install", "--disable-pip-version-check",
                                                  "-r", str(requirements)], capture_output=True, text=True,
                                                 timeout=15 * 60, stdin=subprocess.DEVNULL, creationflags=NO_WINDOW))
    try:
        done = run_pip()
    except (OSError, subprocess.SubprocessError) as e:
        log.write(f"[update] pip failed ({e}): starting with what's installed")
        return
    if done.returncode != 0:
        log.write("[update] pip failed: starting with what's installed. Its last lines:")
        for line in (done.stderr or done.stdout or '').strip().splitlines()[-8:]:
            log.write(f"[update]   {line}")
        return
    Path(stamp).write_text(wanted, encoding="utf-8")
    log.write("[update] requirements installed")


def run():
    """What the scheduled task runs, under pythonw (no window of its own)."""
    if _is_serving():
        return 0
    log = _Log(quiet=True)
    update(log)
    log.write(f"[autostart] Starting app.py in the background (port {PORT}).")
    proc = subprocess.Popen(
        [str(VENV_PYTHON), "-u", "app.py"], cwd=BASE_DIR, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        encoding="utf-8", errors="replace", env={**os.environ, "PYTHONIOENCODING": "utf-8"},
        # A hidden console rather than none: the app's netstat / taskkill calls use it instead of flashing windows
        creationflags=NO_WINDOW)
    for line in proc.stdout:
        log.write(line)
    code = proc.wait()
    log.write(f"[autostart] app.py exited (code {code}).")
    return code


def _task_xml():
    user = escape(f"{os.environ['USERDOMAIN']}\\{os.environ['USERNAME']}")
    # The re-check is a time trigger of its own (a logon trigger's repeats only begin at the next logon), so it's
    # watching from the moment of install. Priority 5 = normal (Task Scheduler's default 7 is below normal).
    return f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Updates the Warehouse app from GitHub's {BRANCH} branch and starts it in the background at logon, and brings it back within {RECHECK_MINUTES} minute(s) if it stops. Managed by {escape(str(BASE_DIR / "autostart.py"))}.</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>{user}</UserId>
    </LogonTrigger>
    <TimeTrigger>
      <Enabled>true</Enabled>
      <StartBoundary>{datetime.now():%Y-%m-%dT%H:%M:%S}</StartBoundary>
      <Repetition>
        <Interval>PT{RECHECK_MINUTES}M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>{user}</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>true</StartWhenAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>5</Priority>
    <Enabled>true</Enabled>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{escape(str(VENV_PYTHONW))}</Command>
      <Arguments>"{escape(str(BASE_DIR / "autostart.py"))}" run</Arguments>
      <WorkingDirectory>{escape(str(BASE_DIR))}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""


def _schtasks(*args, check=True):
    result = subprocess.run(["schtasks", *args], capture_output=True, text=True, stdin=subprocess.DEVNULL)
    if check and result.returncode != 0:
        sys.exit(f"schtasks {args[0]} failed: {(result.stderr or result.stdout).strip()}")
    return result


def _task_status():
    """'Ready' (on), 'Running' (running the app now), 'Disabled' (paused by stop), or None (not set up)."""
    result = _schtasks("/Query", "/TN", TASK_NAME, "/FO", "CSV", "/NH", check=False)
    if result.returncode != 0:
        return None
    return next(csv.reader(result.stdout.strip().splitlines()))[-1]


def _pids_on_port(port):
    """The processes listening on the port (Windows netstat; a listening socket's foreign address ends in :0)."""
    out = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True, creationflags=NO_WINDOW).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[1].endswith(f":{port}") and parts[2].endswith(":0") and parts[4].isdigit():
            pids.add(int(parts[4]))
    return sorted(pids - {0})


def _stop_app():
    """Stops whatever is serving on the app's port and waits for it to be free. Returns whether it did."""
    for pid in _pids_on_port(PORT):
        subprocess.run(["taskkill", "/PID", str(pid), "/F", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    for _ in range(40):
        if not _is_serving():
            break
        time.sleep(0.25)
    else:
        return False
    for _ in range(40):  # let the task's `run` notice and end, or /Run is ignored (IgnoreNew)
        if _task_status() != "Running":
            break
        time.sleep(0.25)
    return True


def _start():
    _schtasks("/Run", "/TN", TASK_NAME)
    for _ in range(120):  # updating (and pip, when requirements changed) comes first
        if _is_serving():
            print(f"The app is running in the background: http://127.0.0.1:{PORT}")
            return
        time.sleep(1)
    print(f"Started, but the app isn't answering yet: see {LOG_PATH}")


def _require_task():
    if _task_status() is None:
        sys.exit("Auto-start isn't set up yet: run  venv\\Scripts\\python autostart.py install")


def install():
    if not VENV_PYTHONW.exists():
        sys.exit(f"{VENV_PYTHONW} not found: set up the venv first (see README).")
    with tempfile.NamedTemporaryFile("w", encoding="utf-16", suffix=".xml", delete=False) as xml_file:
        xml_file.write(_task_xml())
    try:
        _schtasks("/Create", "/TN", TASK_NAME, "/XML", xml_file.name, "/F")
    finally:
        os.unlink(xml_file.name)
    print(f'Scheduled task "{TASK_NAME}" is set up: the app now updates and starts by itself when you log in to Windows.')
    if _is_serving():
        print("The app is already running (e.g. from Start_WarehouseApp.bat), so it's been left alone.\n"
              "The background one takes over within a minute of that one closing, or now with:\n"
              "    venv\\Scripts\\python autostart.py restart")
    else:
        _start()


def status():
    state = _task_status()
    for _ in range(12):  # the per-minute check shows as Running for the moment it takes
        if state != "Running":
            break
        time.sleep(0.25)
        state = _task_status()
    print("Auto-start: " + {
        None: "not set up (run: venv\\Scripts\\python autostart.py install)",
        "Running": "on, running the app in the background",
        "Ready": "on",
        "Disabled": "paused by stop (run: venv\\Scripts\\python autostart.py restart)",
    }.get(state, state))
    try:
        branch = _git(BASE_DIR, "rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
        head = _git(BASE_DIR, "log", "-1", "--format=%h %ad %s", "--date=short").stdout.strip()
        print(f"Code: {branch} @ {head}" + ("" if branch == BRANCH else f"  (not {BRANCH}: auto-update leaves it as it is)"))
    except (OSError, subprocess.SubprocessError):
        print("Code: git isn't available")
    if _is_serving():
        print(f"App: running on http://127.0.0.1:{PORT}" + ("" if state == "Running" else " (started by hand)"))
    else:
        print("App: not running")
    print(f"Log: {LOG_PATH}")


def update_now():
    if _is_serving():
        print("The app is running: the new code is used from its next start (autostart.py restart).")
    print("Update: " + update(_Log()))


def restart():
    _require_task()
    _schtasks("/Change", "/TN", TASK_NAME, "/ENABLE")
    if not _stop_app():
        sys.exit(f"Couldn't stop whatever is using port {PORT}: is it something other than this app?")
    _start()  # the task's run updates before it starts the app


def stop():
    if _task_status() not in (None, "Disabled"):
        _schtasks("/Change", "/TN", TASK_NAME, "/DISABLE")
    if not _stop_app():
        sys.exit(f"Couldn't stop whatever is using port {PORT}: is it something other than this app?")
    print("App stopped, and auto-start paused until: venv\\Scripts\\python autostart.py restart")


def uninstall():
    if _task_status() is not None:
        _schtasks("/Delete", "/TN", TASK_NAME, "/F")
    _stop_app()
    print(f'App stopped and scheduled task "{TASK_NAME}" removed.')


COMMANDS = {"run": run, "install": install, "status": status, "update": update_now, "restart": restart,
            "stop": stop, "uninstall": uninstall}

if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else None
    if command not in COMMANDS:
        sys.exit(__doc__)
    sys.exit(COMMANDS[command]())
