import os
import sys
import zipfile
import shutil
import xml.etree.ElementTree as ET
from lxml import etree
import traceback 
import tempfile




def remove_invalid_character(xml_content):
    # Remove any invalid characters from the XML
    root = ET.fromstring(xml_content)
    for elem in root.iter():
        if elem.text:
            elem.text = ''.join(c for c in elem.text if c.isprintable())
        if elem.tail:
            elem.tail = ''.join(c for c in elem.tail if c.isprintable())
    return ET.tostring(root)


def auto_fix(xml_content):
    parser = etree.XMLParser(recover=True)
    root = etree.fromstring(xml_content, parser)
    return etree.tostring(root)


def fix_word_xml_errors(file_path):
    tempdir = tempfile.mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')

        # Open the Word document file as a zip file
        with zipfile.ZipFile(file_path, 'r') as zipread:
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename == "word/document.xml":
                        # Extract the contents of the word/document.xml file
                        xml_content = zipread.read('word/document.xml')

                        # result_string = remove_invalid_character(xml_content)
                        data = auto_fix(xml_content)
                    else:
                        data = zipread.read(item.filename)
                    zipwrite.writestr(item, data)
        shutil.move(tempname, file_path)
    except Exception:
        traceback.print_exc()
    finally:
        shutil.rmtree(tempdir)


if __name__ == '__main__':
    fix_word_xml_errors(sys.argv[1])
