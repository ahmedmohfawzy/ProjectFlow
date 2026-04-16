
import zipfile
import xml.etree.ElementTree as ET

def peek_excel(file_path):
    try:
        with zipfile.ZipFile(file_path, 'r') as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    for t in tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
                        shared_strings.append(t.text)
            
            if 'xl/worksheets/sheet1.xml' in z.namelist():
                with z.open('xl/worksheets/sheet1.xml') as f:
                    tree = ET.parse(f)
                    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    rows = tree.findall('.//ns:row', ns)
                    for i, row in enumerate(rows[:10]):
                        cells = []
                        for c in row.findall('ns:c', ns):
                            v = c.find('ns:v', ns)
                            t = c.get('t')
                            if t == 's' and v is not None:
                                idx = int(v.text)
                                cells.append(shared_strings[idx] if idx < len(shared_strings) else f"str_{idx}")
                            elif v is not None:
                                cells.append(v.text)
                            else:
                                cells.append("")
                        print(f"Row {i+1}: {cells}")
    except Exception as e:
        print(f"Error: {e}")

peek_excel("/Users/ahmedfawzy/Documents/MS project/TLD F&O and RED365 Project Plan.xlsx")
