"""Export only the public classroom website, without developer or Git files."""
import shutil
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
DEST=ROOT/'build/site'
if not DEST.resolve().is_relative_to(ROOT/'build'):
    raise RuntimeError('Export destination must remain inside the project build directory.')
DEST.mkdir(parents=True,exist_ok=True)
for name in ('index.html','.nojekyll'):
    shutil.copy2(ROOT/name,DEST/name)
for name in ('assets','demos','materials'):
    shutil.copytree(ROOT/name,DEST/name,dirs_exist_ok=True)
print(f'Exported classroom website to {DEST}')
