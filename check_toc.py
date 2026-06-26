import zipfile
with zipfile.ZipFile('Manual_Sistema_Costos_Culinary.docx') as z:
    xml = z.read('word/document.xml').decode('utf-8')

checks = {
    'fldChar begin':    'fldCharType="begin"' in xml,
    'instrText':        '<w:instrText' in xml,
    'TOC instruction':  ' TOC ' in xml,
    'fldChar separate': 'fldCharType="separate"' in xml,
    'fldChar end':      'fldCharType="end"' in xml,
    'Heading1 style':   'Heading1' in xml,
    'Heading2 style':   'Heading2' in xml,
}
for k, v in checks.items():
    print(f'  {k}: {v}')

all_ok = all(checks.values())
print()
print('TOC structure OK:', all_ok)
