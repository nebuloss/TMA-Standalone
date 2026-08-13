import hashlib, os, pathlib, uuid
import re
from xml.sax.saxutils import escape

payload = pathlib.Path(os.environ['TMA_PAYLOAD']).resolve()
output = pathlib.Path(os.environ['TMA_WXS'])
product_version = os.environ.get('TMA_VERSION', '2.0.0')
if not re.fullmatch(r'\d+\.\d+\.\d+', product_version):
    raise ValueError(f'Invalid MSI product version: {product_version}')
files = sorted(p for p in payload.rglob('*') if p.is_file())
dirs = sorted({p.parent.relative_to(payload) for p in files if p.parent != payload},
              key=lambda p: (len(p.parts), str(p)))

def wix_id(prefix, value):
    value = str(value).replace('\\', '/')
    safe = ''.join(c if c.isalnum() or c in '._' else '_' for c in value)
    suffix = hashlib.sha256(value.lower().encode()).hexdigest()[:12]
    return f'{prefix}_{safe[:72-len(prefix)-len(suffix)-2]}_{suffix}'

def stable_guid(value):
    return str(uuid.UUID(bytes_le=hashlib.md5(str(value).lower().encode()).digest())).upper()

ids = {d: wix_id('Dir', d) for d in dirs}
components = []
lines = ['<?xml version="1.0" encoding="utf-8"?>',
 '<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">',
 f'<Product Id="*" Name="TMA autonome" Language="1036" Version="{product_version}" '
 'Manufacturer="TMA Standalone Contributors" UpgradeCode="{4E2C715F-1D33-4769-A56B-4954AD128700}">',
 '<Package InstallerVersion="500" Compressed="yes" InstallScope="perMachine" />',
 '<Media Id="1" Cabinet="payload.cab" EmbedCab="yes" />',
 '<MajorUpgrade AllowSameVersionUpgrades="yes" DowngradeErrorMessage="Une version plus récente de TMA autonome est déjà installée." />',
 '<Directory Id="TARGETDIR" Name="SourceDir"><Directory Id="ProgramFiles64Folder">',
 '<Directory Id="INSTALLFOLDER" Name="TMA-Standalone">']

def emit_files(parent):
    for f in files:
        rel = f.relative_to(payload)
        if rel.parent != parent:
            continue
        component = wix_id('Cmp', rel)
        components.append(component)
        lines.append(f'<Component Id="{component}" Guid="{{{stable_guid(rel)}}}">'
                     f'<File Id="{wix_id("File", rel)}" Source="{escape(str(f))}" KeyPath="yes" /></Component>')

def emit(parent):
    emit_files(parent)
    for d in dirs:
        if d.parent == parent:
            lines.append(f'<Directory Id="{ids[d]}" Name="{escape(d.name)}">')
            emit(d)
            lines.append('</Directory>')
emit(pathlib.Path('.'))

registry = [
 ('ComRegistration', 'A2455011-95CC-42D0-A8AB-8E23C52AF70D', [
  ('Software\\Classes\\TmaCleanRoom.Connect', '', 'TMA autonome Outlook Add-in', 'yes'),
  ('Software\\Classes\\TmaCleanRoom.Connect\\CLSID', '', '{8F5373B8-4973-4E58-A69E-CB57AA22691C}', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', '', 'mscoree.dll', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', 'ThreadingModel', 'Both', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', 'Class', 'TmaCleanRoom.Addin', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', 'Assembly', 'TmaCleanRoom.Addin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', 'RuntimeVersion', 'v4.0.30319', 'no'),
  ('Software\\Classes\\CLSID\\{8F5373B8-4973-4E58-A69E-CB57AA22691C}\\InprocServer32', 'CodeBase', 'file:///[INSTALLFOLDER]TmaCleanRoom.Addin.dll', 'no')]),
 ('OutlookRegistration', 'B4B0DD43-B355-4B1C-B931-52A7768B5F08', [
  ('Software\\Microsoft\\Office\\Outlook\\Addins\\TmaCleanRoom.Connect', 'FriendlyName', 'TMA autonome', 'yes'),
  ('Software\\Microsoft\\Office\\Outlook\\Addins\\TmaCleanRoom.Connect', 'Description', 'Complément Outlook autonome pour les réunions Microsoft Teams', 'no'),
  ('Software\\Microsoft\\Office\\Outlook\\Addins\\TmaCleanRoom.Connect', 'LoadBehavior', '#3', 'no')]),
 ('Resiliency', 'D90EF3B7-708A-43CD-AC59-F86B38DF50E6', [
  ('Software\\Microsoft\\Office\\16.0\\Outlook\\Resiliency\\DoNotDisableAddinList', 'TmaCleanRoom.Connect', '#1', 'yes')])]
for cid, guid, values in registry:
    components.append(cid)
    lines.append(f'<Component Id="{cid}" Guid="{{{guid}}}">')
    for index, (key, name, value, keypath) in enumerate(values):
        name_attr = f' Name="{escape(name)}"' if name else ''
        value_type = 'integer' if value.startswith('#') else 'string'
        clean_value = value[1:] if value_type == 'integer' else value
        lines.append(f'<RegistryValue Id="Reg_{cid}_{index}" Root="HKCU" Key="{escape(key)}"{name_attr} '
                     f'Type="{value_type}" Value="{escape(clean_value)}" KeyPath="{keypath}" />')
    lines.append('</Component>')
lines += ['</Directory></Directory></Directory>', '<Feature Id="MainFeature" Title="TMA autonome" Level="1">']
lines += [f'<ComponentRef Id="{component}" />' for component in components]
lines += ['</Feature>', '</Product>', '</Wix>']
output.write_text('\n'.join(lines) + '\n', encoding='utf-8')
