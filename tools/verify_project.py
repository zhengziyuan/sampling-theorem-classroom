"""Validate the classroom package using only the Python standard library."""
import hashlib
import json
import re
import sys
import wave
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
errors = []
checked_links = 0

class Links(HTMLParser):
    def handle_starttag(self, tag, attrs):
        for key, value in attrs:
            if value and key in {'href','src'}:
                check_link(self.source, value)

def check_link(source, link):
    global checked_links
    parsed = urlsplit(link)
    if parsed.scheme or parsed.netloc or not parsed.path:
        return
    target = (source.parent / unquote(parsed.path)).resolve()
    if not target.is_relative_to(ROOT):
        errors.append(f'Link escapes project: {source.relative_to(ROOT)}')
    elif not target.exists():
        errors.append(f'Missing link target: {source.relative_to(ROOT)} -> {link}')
    checked_links += 1

def verify():
    manifest = json.loads((ROOT / 'materials/manifest.json').read_text(encoding='utf-8'))
    for item in manifest['preserved_files']:
        path = ROOT / item['path']
        if not path.is_file():
            errors.append(f'Missing original file: {item["path"]}')
            continue
        data = path.read_bytes()
        if len(data) != item['bytes'] or hashlib.sha256(data).hexdigest() != item['sha256']:
            errors.append(f'Original file changed: {item["path"]}')
    for path in ROOT.glob('materials/**/*'):
        if path.suffix.lower() not in {'.pptx','.pptm','.docx'}:
            continue
        try:
            with zipfile.ZipFile(path) as z:
                bad = z.testzip()
                if bad:
                    errors.append(f'Corrupt Office part: {path.name}: {bad}')
                names = set(z.namelist())
                for part in names:
                    if not part.endswith('.rels'):
                        continue
                    for rel in ET.fromstring(z.read(part)):
                        if rel.attrib.get('TargetMode') == 'External':
                            target=rel.attrib.get('Target','')
                            if re.match(r'^(file:|[A-Za-z]:[\\/])',target):
                                errors.append(f'Local-machine Office link: {path.name}')
                if path.suffix.lower() in {'.pptx','.pptm'}:
                    expected = next(x for x in manifest['slides'] if x['path'] == path.relative_to(ROOT).as_posix())
                    count = sum(bool(re.fullmatch(r'ppt/slides/slide\d+\.xml',p)) for p in names)
                    if count != expected['slides']:
                        errors.append(f'Slide count changed: {path.name}')
        except (zipfile.BadZipFile,ET.ParseError) as exc:
            errors.append(f'Invalid Office file: {path.name}: {exc}')
    pages = [ROOT / 'index.html', *sorted((ROOT/'demos').glob('*.html'))]
    for path in pages:
        parser=Links()
        parser.source=path
        parser.feed(path.read_text(encoding='utf-8'))
    for path in [ROOT/'README.md',ROOT/'NOTICE.md',*sorted((ROOT/'docs').glob('*.md'))]:
        for target in re.findall(r'\]\(([^\s)]+)\)',path.read_text(encoding='utf-8')):
            check_link(path,target)
    audio_manifest=json.loads((ROOT/'assets/audio/manifest.json').read_text(encoding='utf-8'))
    for clip in audio_manifest['clips']:
        path=ROOT/'assets/audio'/clip['file']
        if not path.is_file():
            errors.append(f'Missing audio: {clip["file"]}')
            continue
        with wave.open(str(path)) as stream:
            duration=stream.getnframes()/stream.getframerate()
            if stream.getframerate()!=44100 or stream.getnchannels()!=1 or not 10.4<=duration<=10.6:
                errors.append(f'Unexpected WAV parameters: {path.name}')
    for image in ('spectra_image','waveform_image'):
        if not (ROOT/'assets/audio'/audio_manifest[image]).is_file():
            errors.append(f'Missing audio figure: {image}')
    for path in ROOT.rglob('*'):
        if not path.is_file() or any(p in {'.git','build','__pycache__'} for p in path.relative_to(ROOT).parts):
            continue
        if path.suffix.lower() in {'.html','.css','.json','.py','.ps1','.md','.bas','.yml'}:
            content=path.read_text(encoding='utf-8-sig')
            if re.search(r'[A-Za-z]:[\\/]+Users[\\/]',content,re.I):
                errors.append(f'Personal absolute path: {path.relative_to(ROOT)}')
            if re.search(r'gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{40,}|sk-proj-[A-Za-z0-9_-]{25,}',content):
                errors.append(f'Possible credential: {path.relative_to(ROOT)}')
    result={'preserved_files':len(manifest['preserved_files']),'slide_decks':len(manifest['slides']),
            'web_pages':len(pages),'checked_local_links':checked_links,'audio_clips':len(audio_manifest['clips']),
            'errors':errors}
    print(json.dumps(result,ensure_ascii=False,indent=2))
    return not errors

if __name__=='__main__':
    sys.exit(0 if verify() else 1)
