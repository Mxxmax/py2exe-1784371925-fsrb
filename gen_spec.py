#!/usr/bin/env python3
"""Generate PyInstaller .spec file with Tree() injection for openpyxl/et_xmlfile."""
import sys
from pathlib import Path
import openpyxl, et_xmlfile

def gen_spec(source_py: str, app_name: str, spec_path: str):
    op = Path(openpyxl.__file__).resolve().parent
    et = Path(et_xmlfile.__file__).resolve().parent
    parts = [
        "a = Analysis(",
        f"    [r'{source_py}'],",
        "    pathex=[],",
        "    binaries=[], datas=[],",
        "    hiddenimports=['openpyxl','et_xmlfile'],",
        "    hookspath=[], hooksconfig={}, runtime_hooks=[],",
        "    excludes=['IPython','matplotlib','numpy','pandas','PIL','Pillow','scipy','sklearn','tensorflow','torch'],",
        "    win_no_prefer_redirects=False, win_private_assemblies=False,",
        "    cipher=None, noarchive=False,",
        ")",
        "pyz = PYZ(a.pure, a.zipped_data, cipher=None)",
        "exe = EXE(",
        "    pyz, a.scripts,",
        f"    Tree(r'{op}', prefix='openpyxl'),",
        f"    Tree(r'{et}', prefix='et_xmlfile'),",
        "    a.binaries, a.zipfiles, a.datas,",
        f"    name='{app_name}',",
        "    debug=False, bootloader_ignore_signals=False,",
        "    strip=False, upx=True, console=True,",
        "    disable_windowed_traceback=False, argv_emulation=False,",
        "    target_arch=None, codesign_identity=None, entitlements_file=None,",
        ")",
    ]
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write('# -*- mode: python ; coding: utf-8 -*-\n')
        f.write('from pathlib import Path\n')
        f.write('\n'.join(parts))
        f.write('\n')
    print(f'Spec generated: {spec_path}')

if __name__ == '__main__':
    gen_spec(sys.argv[1], sys.argv[2], sys.argv[3])
