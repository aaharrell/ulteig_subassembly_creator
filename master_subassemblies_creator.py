import os
from openpyxl import load_workbook
import pandas as pd


# Linking helper function
def make_link(drawing_path, drawing_name):
    joined_path = os.path.join(drawing_path, drawing_name)
    return '=HYPERLINK("{}", "{}")'.format(joined_path, drawing_name)

def main():

    # Get path to the subassemblies folder
    print("Ensure the subassembly drawings are in the final desired directory to ensure correct linking.")
    print("Go the folder with your subassemblies using file explorer. Copy the path to this folder using the file explorer address bar.")
    print("NOTE: This folder should ONLY contain the subassembly drawing files.\n")
    drawing_path = input("Paste the path here and press Enter: ")
    print("Processing...\n")

    # Get path as string literal
    drawing_path = r'{}'.format(drawing_path)

    # Get list of drawings
    try:
        drawing_list = os.listdir(drawing_path)
    except:
        print("ERROR: Drawing directory not found; please verify the correct file location of the drawings.\n")
        main()
        return

    # Create dataframe using Pandas and list of drawings
    col1 = 'Drawing Name'
    col2 = 'Major Subassembly Group'
    col3 = 'Minor Subassembly Group'
    df = pd.DataFrame(drawing_list, index=None, columns=[col1], dtype=None, copy=None)
    df[col2] = "-"
    df[col3] = "-"

    # Create links
    for i in range(len(df.index)):
        link = make_link(drawing_path, drawing_list[i])
        df.at[i, col1] = link
        df.at[i, col2] = drawing_list[i].split("-")[0]
        df.at[i, col3] = drawing_list[i].split("-")[1]

    # Create spreadsheet
    spreadsheet_name = "Master Subassemblies List.xlsx"
    df.to_excel(spreadsheet_name, index=False)

    # Format spreadsheet
    wb = load_workbook(spreadsheet_name)
    ws = wb['Sheet1']
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    wb.save(spreadsheet_name)

    userIn = "-"
    while userIn != "":
        userIn = input("Operation success; check the folder containing this application for the spreadsheet.\nPress Enter to exit: ")

main()