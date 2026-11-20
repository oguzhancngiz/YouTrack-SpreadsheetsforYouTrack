# Copyright 2026 Codeo
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

"""Captures the Marketplace screenshots from the development preview with headless Microsoft Edge (or Chrome).

    python serve.py 8766          # in another terminal
    python docs/capture.py        # writes docs/marketplace/screenshots/*.png
Not part of the release ZIP.
"""
import os
import shutil
import subprocess
import sys
import tempfile

BASE = 'http://127.0.0.1:%s/widgets/sheet/' % (sys.argv[1] if len(sys.argv) > 1 else '8766')
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'marketplace', 'screenshots')
BROWSERS = [r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', 'google-chrome', 'chromium']
SHOTS = [  # all 1280 x 800 (one aspect ratio), light theme (the YouTrack default)
    ('screenshot-1-spreadsheet.png', 'index.html?user=alice&demo&height=700&select=F2:F9'),
    ('screenshot-2-filter.png', 'index.html?user=alice&demo&height=700&action=filter'),
    ('screenshot-3-formulas.png', 'index.html?user=alice&demo&height=700&select=K7&action=edit'),
    ('screenshot-4-test-matrix.png', 'index.html?user=alice&demo&height=700&sheet=2&select=B2:E5'),
    ('screenshot-5-large-window.png', 'window.html?user=alice&demo&select=G4'),
    ('screenshot-6-conditional-formatting.png', 'window.html?user=alice&demo&action=cf'),
    ('screenshot-7-history.png', 'window.html?user=alice&demo&action=history'),
    ('screenshot-8-solver.png', 'window.html?user=alice&demo&sheet=3&select=B6&action=solver'),
    ('screenshot-9-solver-results.png', 'window.html?user=alice&demo&sheet=3&select=B6&action=solve'),
    ('extra/dark-theme.png', 'window.html?user=alice&demo&theme=dark&select=G4'),
]


def browser():
    for b in BROWSERS:
        if os.path.isfile(b) or shutil.which(b):
            return b
    raise SystemExit('No Chrome or Edge found')


def main():
    os.makedirs(OUT, exist_ok=True)
    exe = browser()
    w, h = 1280, 800
    for name, url in SHOTS:
        profile = tempfile.mkdtemp(prefix='issue-sheets-shot-')
        path = os.path.join(OUT, name)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        cmd = [exe, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--force-device-scale-factor=1',
               '--no-first-run', '--user-data-dir=' + profile, '--window-size=%d,%d' % (w, h),
               '--virtual-time-budget=6000', '--screenshot=' + path, BASE + url]
        subprocess.run(cmd, check=False, timeout=90, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        shutil.rmtree(profile, ignore_errors=True)
        print(('ok  ' if os.path.isfile(path) else 'FAIL') + ' ' + name)


if __name__ == '__main__':
    main()
