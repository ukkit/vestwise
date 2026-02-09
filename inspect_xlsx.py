import zipfile
import xml.etree.ElementTree as ET

# XLSX files are zip archives
with zipfile.ZipFile('BenefitHistory.xlsx', 'r') as zip_ref:
    # Read shared strings
    shared_strings_xml = zip_ref.read('xl/sharedStrings.xml').decode('utf-8')
    root_ss = ET.fromstring(shared_strings_xml)
    shared_strings = []
    for si in root_ss.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
        t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
        if t is not None:
            shared_strings.append(t.text)

    print(f'Shared strings count: {len(shared_strings)}')

    # Get sheet1.xml (ESPP)
    sheet1_xml = zip_ref.read('xl/worksheets/sheet1.xml').decode('utf-8')
    print('\n--- ESPP Sheet Columns ---')
    root1 = ET.fromstring(sheet1_xml)

    rows = root1.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
    if rows:
        header_row = rows[0]
        cells = header_row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
        print(f'Total columns: {len(cells)}')
        for i, cell in enumerate(cells):
            v = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v is not None and v.text:
                try:
                    idx = int(v.text)
                    if idx < len(shared_strings):
                        print(f'  {i+1:2d}. {shared_strings[idx]}')
                except:
                    pass

    # Get sheet2.xml (Restricted Stock)
    sheet2_xml = zip_ref.read('xl/worksheets/sheet2.xml').decode('utf-8')
    print('\n--- Restricted Stock Sheet Columns ---')
    root2 = ET.fromstring(sheet2_xml)

    rows = root2.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')
    if rows:
        header_row = rows[0]
        cells = header_row.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')
        print(f'Total columns: {len(cells)}')
        for i, cell in enumerate(cells):
            v = cell.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
            if v is not None and v.text:
                try:
                    idx = int(v.text)
                    if idx < len(shared_strings):
                        print(f'  {i+1:2d}. {shared_strings[idx]}')
                except:
                    pass
