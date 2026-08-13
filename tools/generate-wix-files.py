import hashlib, os, pathlib, uuid
from xml.sax.saxutils import escape

payload = pathlib.Path(os.environ['TMA_PAYLOAD']).resolve()
output = pathlib.Path(os.environ['TMA_WXS'])
files = sorted(p for p in payload.rglob('*') if p.is_file())
dirs = sorted({p.parent.relative_to(payload) for p in files if p.parent != payload}, key=lambda p: (len(p.parts), str(p)))

def wix_id(prefix, value):
    value = str(value).replace('\\', '/')
    safe = ''.join(c if c.isalnum() or c in '._' else '_' for c in value)
    suffix = hashlib.sha256(value.lower().encode()).hexdigest()[:12]
    return f'{prefix}_{safe[:72-len(prefix)-len(suffix)-2]}_{suffix}'

def stable_guid(value):
    digest = bytearray(hashlib.md5(str(value).lower().encode()).digest())
    return str(uuid.UUID(bytes_le=bytes(digest)))

ids = {d: wix_id('Dir', d) for d in dirs}
lines = ['<?xml version="1.0" encoding="utf-8"?>',
 '<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs"><Fragment><DirectoryRef Id="INSTALLFOLDER">']
def emit(parent, indent):
    for d in dirs:
        if d.parent == parent:
            lines.append(' ' * indent + f'<Directory Id="{ids[d]}" Name="{escape(d.name)}">')
            emit(d, indent + 2)
            lines.append(' ' * indent + '</Directory>')
emit(pathlib.Path('.'), 4)
lines.append('</DirectoryRef></Fragment><Fragment><ComponentGroup Id="TmaFiles">')
for f in files:
    rel = f.relative_to(payload)
    directory = 'INSTALLFOLDER' if rel.parent == pathlib.Path('.') else ids[rel.parent]
    lines.append(f'<Component Id="{wix_id("Cmp", rel)}" Guid="{stable_guid(rel)}" Directory="{directory}">'
                 f'<File Id="{wix_id("File", rel)}" Source="{escape(str(f))}" KeyPath="yes" /></Component>')
lines.append('</ComponentGroup></Fragment></Wix>')
output.write_text('\n'.join(lines) + '\n', encoding='utf-8')
