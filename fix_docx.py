import zipfile, shutil, os, re

src = 'Manual_Sistema_Costos_Culinary.docx'
tmp = '_manual_tmp'

if os.path.exists(tmp):
    shutil.rmtree(tmp)
with zipfile.ZipFile(src, 'r') as z:
    z.extractall(tmp)

sfile = os.path.join(tmp, 'word', 'settings.xml')
with open(sfile, 'rb') as f:
    content = f.read().decode('utf-8')

# Add w:percent="100" to w:zoom if missing
if 'w:zoom' in content and 'w:percent' not in content:
    content = content.replace('<w:zoom/>', '<w:zoom w:percent="100"/>')
    content = re.sub(r'<w:zoom\s+(?!w:percent)', '<w:zoom w:percent="100" ', content)

with open(sfile, 'w', encoding='utf-8') as f:
    f.write(content)

os.remove(src)
with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as zout:
    for root, dirs, files in os.walk(tmp):
        for fname in files:
            fpath = os.path.join(root, fname)
            arcname = os.path.relpath(fpath, tmp)
            zout.write(fpath, arcname)

shutil.rmtree(tmp)
print('Done: settings.xml fixed and repacked OK')
