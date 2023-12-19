from openpyxl import load_workbook
import os
import pandas as pd

# User instructions + get path to subassembly folder
def get_subassembly_path():
    print("READ ME:")
    print("\t1. Ensure the subassembly drawings are in the final desired\n\t   directory to ensure correct linking.")
    print("\t\tNOTE: Make sure there are no files open. If the subassemblies\n\t\t      are in OneDrive, make sure all of the files have downloaded.\n")
    print("\t2. Go the folder with your subassemblies using file explorer.\n\t   Copy the path to this folder using the file explorer address bar.")
    print("\t\tNOTE: This folder should ONLY contain the subassembly drawing files\n\t\t      and optionally the Excel file with drawing information. See (3) below.\n")
    print("\t3. If you have included the ProjectWise drawing information spreadsheet,\n\t   ensure the spreadsheet name begins with \"_\"; e.g. \"_Subassembly Dwg Info.xlsx\"\n")
    userIn = input("Paste the folder path here and press Enter: ")

    # Get path as string literal
    userIn = r'{}'.format(userIn)
    print("Processing...\n")
    return(userIn)

# Linking helper function
def make_link(drawing_path, drawing_name):
    joined_path = os.path.join(drawing_path, drawing_name)
    return '=HYPERLINK("{}", "{}")'.format(joined_path, drawing_name)

# Get the subassembly drawing info and create a dataframe
def create_subassy_info_df(drawing_path, drawing_list):
    if drawing_list[-1].split(".")[1] == "xlsx":
        df_subassy_info = pd.read_excel(os.path.join(drawing_path, drawing_list[-1]))
        return(df_subassy_info)
    
# Create and format spreadsheet
def create_spreadsheet(spreadsheet_name, df):
    # Create spreadsheet
    df.to_excel(spreadsheet_name, index=False)

    # Format spreadsheet
    wb = load_workbook(spreadsheet_name)
    ws = wb['Sheet1']
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 50
    ws.column_dimensions['F'].width = 50
    ws.column_dimensions['G'].width = 10
    wb.save(spreadsheet_name)
    return(wb)

def main():

    # Get path to the subassemblies folder
    drawing_path = get_subassembly_path()

    # Get list of drawings
    try:
        drawing_list = os.listdir(drawing_path)
    except:
        print("ERROR: Drawing directory not found; please verify the correct file location of the drawings.\n")
        main()
        return

    # Create primary dataframe using Pandas and subassembly drawing list
    col1 = "Drawing Name (Linked)"
    col2 = "Major Subassembly Group"
    col3 = "Minor Subassembly Group"
    col4 = "Drawing Title"
    col5 = "Description 1"
    col6 = "Description 2"
    col7 = "Rev"

    df_main = pd.DataFrame(drawing_list[:-2], index=None, columns=[col1], dtype=None, copy=None)
    df_main[col2] = "-"
    df_main[col3] = "-"
    df_main[col4] = "-"
    df_main[col5] = "-"
    df_main[col6] = "-"
    df_main[col7] = "-"

    # Create second dataframe with subassembly drawing information
    df_subassy_info = create_subassy_info_df(drawing_path, drawing_list)

    # Create links for column 1 and add other columns
    for i in range(len(df_main.index)):
        link = make_link(drawing_path, drawing_list[i])
        df_main.at[i, col1] = link
        df_main.at[i, col2] = drawing_list[i].split("-")[0]
        df_main.at[i, col3] = drawing_list[i].split("-")[1]
        df_main.at[i, col4] = df_subassy_info.at[i, "dwgtitle"]
        df_main.at[i, col5] = df_subassy_info.at[i, "dwgtitle3"]
        df_main.at[i, col6] = df_subassy_info.at[i, "dwgtitle4"]
        df_main.at[i, col7] = df_subassy_info.at[i, "revision"]

    create_spreadsheet("Master Subassemblies List.xlsx", df_main)

    userIn = "-"
    while userIn != "":
        userIn = input("Operation success; check the folder containing this application for the spreadsheet.\nPress Enter to exit: ")

main()