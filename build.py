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

"""Builds the release package of Spreadsheets for YouTrack.

    python build.py

1. Copies sheet-core.js (shared by the backend and the widget) to widgets/sheet/.
2. Writes dist/ with only the files YouTrack needs; the development preview (preview.js, test.html, the code between
   @preview-start and @preview-end) is left out.
3. Zips dist/ into issue-sheets.zip, which you upload in YouTrack: Administration > Apps > Add app > Upload ZIP file.
"""
import json
import os
import re
import shutil
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(ROOT, 'dist')
ZIP = os.path.join(ROOT, 'spreadsheets-for-youtrack.zip')

# README.md stays out of the package: it has development URLs, which the YouTrack app review does not allow there
ROOT_FILES = ['manifest.json', 'settings.json', 'entity-extensions.json', 'backend.js', 'sheet-core.js',
              'icon.svg', 'icon-dark.svg', 'LICENSE', 'NOTICE', 'THIRD-PARTY-NOTICES.txt']
WIDGET_FILES = ['index.html', 'window.html', 'app.js', 'app.css', 'ring-variables.css', 'formula.js', 'sync.js',
                'model.js', 'io.js', 'charts.js', 'ui.js', 'dialogs.js', 'solver.js', 'whatif.js', 'sheet-core.js', 'icon.svg']
PREVIEW_RE = re.compile(r'[ \t]*// @preview-start.*?// @preview-end[^\n]*\n', re.S)
FORBIDDEN = ['preview.js', '127.0.0.1', 'localhost']


def main():
    shutil.copyfile(os.path.join(ROOT, 'sheet-core.js'), os.path.join(ROOT, 'widgets', 'sheet', 'sheet-core.js'))
    manifest = json.load(open(os.path.join(ROOT, 'manifest.json'), encoding='utf-8'))
    if os.path.isdir(DIST):
        shutil.rmtree(DIST)
    os.makedirs(os.path.join(DIST, 'widgets', 'sheet'))
    for name in ROOT_FILES:
        shutil.copyfile(os.path.join(ROOT, name), os.path.join(DIST, name))
    for name in WIDGET_FILES:
        src = os.path.join(ROOT, 'widgets', 'sheet', name)
        dst = os.path.join(DIST, 'widgets', 'sheet', name)
        if name.endswith('.js'):
            text = open(src, encoding='utf-8').read()
            text = PREVIEW_RE.sub('', text)
            for word in FORBIDDEN:
                if word in text:
                    raise SystemExit(f'{name} still contains {word!r}')
            open(dst, 'w', encoding='utf-8', newline='\n').write(text)
        else:
            shutil.copyfile(src, dst)
    # every widget path in the manifest must exist
    for w in manifest['widgets']:
        for key in ('indexPath', 'iconPath'):
            path = os.path.join(DIST, 'widgets', w[key])
            if not os.path.isfile(path):
                raise SystemExit(f'Missing {w[key]} for widget {w["key"]}')
    for key in ('icon', 'iconDark'):
        if key in manifest and not os.path.isfile(os.path.join(DIST, manifest[key])):
            raise SystemExit(f'Missing {manifest[key]}')
    if os.path.exists(ZIP):
        os.remove(ZIP)
    with zipfile.ZipFile(ZIP, 'w', zipfile.ZIP_DEFLATED) as z:
        for folder, _, files in os.walk(DIST):
            for f in sorted(files):
                full = os.path.join(folder, f)
                z.write(full, os.path.relpath(full, DIST).replace(os.sep, '/'))
    print(f'Built {os.path.basename(ZIP)}: {manifest["title"]} {manifest["version"]}')


if __name__ == '__main__':
    main()
