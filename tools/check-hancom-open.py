#!/usr/bin/env python3
"""Open generated HWPs in Hancom on an isolated X display; require stable windows.

Requires Linux, xdotool, and a running X server (prefer a dedicated Xvfb).
Pass the location-independent run-hwp.sh wrapper as --executable.
This checks loading and initial layout, not all-page visual fidelity.
"""

import argparse
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import time


def window_titles(pid, env):
    ids = subprocess.run(
        ['xdotool', 'search', '--pid', str(pid)], env=env,
        capture_output=True, text=True, check=False,
    ).stdout.split()
    return [subprocess.run(
        ['xdotool', 'getwindowname', wid], env=env,
        capture_output=True, text=True, check=False,
    ).stdout.strip() for wid in ids]


def check_file(executable, path, env, timeout, settle):
    path = path.resolve(strict=True)
    process = subprocess.Popen(
        [str(executable), str(path)], env=env, start_new_session=True,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    titles = []
    opened = False
    try:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline and process.poll() is None:
            titles = window_titles(process.pid, env)
            if any(path.name in title for title in titles):
                opened = True
                break
            time.sleep(0.1)
        if opened:
            # A title alone passed even when the calendar crashed on first paint.
            end = time.monotonic() + settle
            while time.monotonic() < end and process.poll() is None:
                time.sleep(0.1)
            titles = window_titles(process.pid, env)
            opened = process.poll() is None and any(path.name in title for title in titles)
        return {
            'file': str(path), 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(),
            'opened': opened, 'settle_seconds': settle, 'titles': titles,
            'exit_code_before_cleanup': process.poll(),
        }
    finally:
        # Only terminate the process group launched for this check.
        try:
            os.killpg(process.pid, signal.SIGTERM)
            process.wait(timeout=2)
        except ProcessLookupError:
            pass
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGKILL)
            process.wait()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--executable', type=Path, required=True)
    parser.add_argument('--display', required=True, help='Dedicated X display, e.g. :117')
    parser.add_argument('--timeout', type=float, default=10)
    parser.add_argument('--settle', type=float, default=5)
    parser.add_argument('files', type=Path, nargs='+')
    args = parser.parse_args()
    if args.timeout <= 0 or args.settle <= 0:
        parser.error('timeout and settle must be positive')
    executable = args.executable.resolve(strict=True)
    env = dict(os.environ, DISPLAY=args.display, QT_IM_MODULE='compose')
    failed = False
    for path in args.files:
        result = check_file(executable, path, env, args.timeout, args.settle)
        print(json.dumps(result, ensure_ascii=False), flush=True)
        failed |= not result['opened']
    return int(failed)


if __name__ == '__main__':
    raise SystemExit(main())
