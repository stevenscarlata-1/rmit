# Usage: python3 test/check-xml.py out.docx  — fails if any XML part is malformed
import sys, zipfile, xml.dom.minidom as m
z = zipfile.ZipFile(sys.argv[1]); bad = 0
for n in z.namelist():
    if n.endswith(('.xml', '.rels')):
        try: m.parseString(z.read(n))
        except Exception as e: print('BAD', n, e); bad += 1
print('XML OK' if not bad else f'{bad} bad parts'); sys.exit(1 if bad else 0)
