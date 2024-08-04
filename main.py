import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import sys
from datetime import datetime
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Get the path of the script
bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
path_to_model = os.path.abspath(os.path.join(bundle_dir, 'model.xlsx'))

def get_text(element, tag):
    child = element.find(tag)
    return child.text if child is not None else ''

def main(excel_file, xml_file):
    workbook = load_workbook(excel_file)
    sheet = workbook.active
    tree = ET.parse(xml_file)
    root = tree.getroot()
    row = 3
    for item in root.findall('.//G_ITENS_SOLICITACAO'):
        sheet.cell(row=row, column=2).value = get_text(item, 'DESCRICAO')
        sheet.cell(row=row, column=3).value = get_text(item, 'QTDE')
        row += 1

    # Generate the filename with current date and time
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    new_filename = f'cotação {timestamp}.xlsx'
    workbook.save(new_filename)
    print(f"Saved file as {new_filename}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        # Open a file dialog box to select the XML file
        root = Tk()
        root.withdraw()  # Hide the root window
        xml_file = askopenfilename(filetypes=[("XML files", "*.xml")])
        if not xml_file:
            print("No file selected. Exiting.")
            sys.exit(1)
    else:
        xml_file = sys.argv[1]

    print(f"Processing file: {xml_file}")
    excel_file = path_to_model
    main(excel_file, xml_file)
