"""Claude Code hook, after every file edit: lint the Python file that changed and run the tests, so a broken
reader or label never gets past the edit that broke it. Problems go back to Claude (exit 2); runs no AI itself.

Also runnable by hand:  python .claude/hooks/after_edit.py path/to/file.py
"""
import json
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# Files whose changes can move a label, a code or a courier reference: the tests cover them
WATCHED = ('horizontal_', 'packing_label', 'courier_', 'label_stitcher', 'core_math', 'matrix_engine', 'app.py', 'tests/')


def changed_file():
    if len(sys.argv) > 1:
        return sys.argv[1]
    try:
        return (json.load(sys.stdin).get('tool_input') or {}).get('file_path', '')
    except ValueError:
        return ''


def main():
    path = changed_file()
    if not path.endswith('.py'):
        return 0
    rel = os.path.relpath(os.path.abspath(path), ROOT).replace('\\', '/')
    problems = []
    if shutil.which('ruff'):
        lint = subprocess.run(['ruff', 'check', '--select', 'E9,F', '--quiet', path], cwd=ROOT, capture_output=True, text=True)
        if lint.returncode:
            problems.append(f"ruff found problems in {rel}:\n{lint.stdout or lint.stderr}")
    if any(w in rel for w in WATCHED):
        tests = subprocess.run([sys.executable, '-m', 'pytest', '-q', '-x', 'tests'], cwd=ROOT, capture_output=True,
                               text=True, timeout=170)
        if tests.returncode:
            problems.append("Tests failed after this edit:\n" + "\n".join(tests.stdout.splitlines()[-25:]))
    if problems:
        print("\n\n".join(problems), file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
