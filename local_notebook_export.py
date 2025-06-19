import os
import pprint
import win32com.client as win32
from pathlib import Path
import lxml.etree as ET
from tqdm import tqdm
import time

def validate_with_xsd(xml_content, xsd_path):
    """Validate XML against a local XSD schema"""
    try:
        # Load XSD schema
        with open(xsd_path, 'rb') as f:
            xsd_content = f.read()
        schema = ET.XMLSchema(ET.fromstring(xsd_content))
        
        # Parse XML
        parser = ET.XMLParser(schema=schema)
        xml_doc = ET.fromstring(xml_content, parser)
        return xml_doc
        
    except ET.XMLSchemaError as e:
        print(f"Schema validation error: {e}")
        return None
    except ET.XMLSyntaxError as e:
        print(f"XML syntax error: {e}")
        return None

def sanitize_name(name):
    """Sanitize filenames for Windows"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name

xsd_path = "0336.OneNoteApplication_2013.xsd"  # OneNote 2013 schema xsd
onenote = win32.gencache.EnsureDispatch('OneNote.Application')
hierarchy = ""
hierarchy = onenote.GetHierarchy("", 2) # gets all local notebooks and unfiled notes

root = validate_with_xsd(hierarchy, xsd_path)
if root is not None:
    print("XML successfully validated against schema")
    tree = ET.ElementTree(root)
    tree.write("hierarchy.xml", encoding="utf-8", xml_declaration=True, pretty_print=True)

notebook_tag = '{http://schemas.microsoft.com/office/onenote/2013/onenote}Notebook'

iter_count = 0
iter_count = len([1 for c in root.iter(tag=notebook_tag)])
print(f"{iter_count} notebooks to process")

failed_exports = []

for child in tqdm(root.iter(tag=notebook_tag), total=iter_count):
    notebook_dict = child.attrib
    tqdm.write(f"Exporting: {notebook_dict["name"]}")
    
    notebook_id = notebook_dict['ID']
    notebook_name = notebook_dict['name']
    export_path = (Path(__file__).parent.absolute() / "Backups" / sanitize_name(notebook_name)).with_suffix(".onepkg")
    
    # Skips already exported notebooks
    if Path(export_path).exists():
        continue
    
    onenote.Publish(notebook_id, str(export_path), 1)

    # Wait until the file exists and is stable in size
    max_wait = 300  # seconds
    waited = 0
    last_size = -1
    while waited < max_wait:
        if os.path.exists(export_path):
            size = os.path.getsize(export_path)
            if size == last_size and size > 0:
                break  # File size is stable, assume export is done
            last_size = size
        time.sleep(1)
        waited += 1
    else:
        tqdm.write(f"Warning: Export file {export_path} may not be complete after {max_wait} seconds.")
        if not os.path.exists(export_path):
            failed_exports.append(notebook_name)
            
if len(failed_exports) != 0:
    print("problematic exports:")
    pprint.pprint(failed_exports)