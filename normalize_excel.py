# Project: Normalize Excel
# Description:
# Author: EHCIKNA

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
import os

# Basic colors
red_fill     = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
green_fill   = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
blue_fill    = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
yellow_fill  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
pink_fill    = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

# Extras (optional)
orange_fill  = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
purple_fill  = PatternFill(start_color="800080", end_color="800080", fill_type="solid")
gray_fill    = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

def formatting(ws):
    try:
        center_align=Alignment(horizontal='center',vertical='center',text_rotation=0)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.alignment=center_align
        return True
    except Exception as e:
        print(f">> Function:{formatting.__name__}, Error processing file: {str(e)}")
        return False

def unmerge_fill(ws,issues):
    try:
        for merge_range in list(ws.merged_cells.ranges):
            cell_value = ws[merge_range.coord.split(":")[0]].value  # extract merge cell value
            if cell_value is None:
                issues['merge_empty_cells'].append(merge_range)
            ws.unmerge_cells(str(merge_range))  # unmerge cell range

            # setting same merge value for all cells after unmerge
            for row in ws[merge_range.coord]:
                for cell in row:
                    if cell_value is None:
                        cell.fill = red_fill
                    else:
                        cell.value = cell_value
        return True
    except Exception as e:
        print(f">> Function:{unmerge_fill.__name__}, Error processing file: {str(e)}")
        return False

def processing_excel(*file_info):
    """
    Unmerges all merged cells in the given Excel sheet and fills each
    unmerged cell with the original merged value.

    Parameters
    ----------
    *file_info : tuple
        (input_file_name, sheet_name, output_file_name)
    """
    input_file_name=file_info[0]
    sheet_name=file_info[1]
    output_file_name=file_info[2]
    issues=file_info[3]

    try:
        #open workbook
        if not os.path.exists(input_file_name):
            print(f">> '{input_file_name}' does not exist")
            return False
        wb= load_workbook(input_file_name, read_only=False)

        if sheet_name not in wb.sheetnames:
            print(f">> '{sheet_name}' does not exist. Available sheet names are \"{wb.sheetnames}\"")
            return False
        ws= wb[sheet_name]

        #Unmerge Processing
        unmerge_status=unmerge_fill(ws,issues)
        if not unmerge_status:
            print(f">> Unmerge Failed")
            return False

        #correcting alignment
        align_status=formatting(ws)
        if not align_status:
            print(f">> Alignment Failed")
            return False

        # save workbook
        wb.save(output_file_name)
        wb.close()
        return True
    except PermissionError:
        print(f">> Permission denied. Make sure '{input_file_name}' or '{output_file_name}' is not open in Excel")
        return False
    except Exception as e:
        print(f">> Error processing file: {str(e)}")
        return False


if __name__ == "__main__":
    #configration
    input_file_name= "input.xlsx"
    sheet_name= "Node Version Planner"
    # file_datetime=datetime.now().strftime("%d-%m-%Y_%H-%M-%S") #datetime format is dd-mm-yy_hour-min-second
    # output_file_name= "unmerged_output_" + str(file_datetime) + ".xlsx"
    output_file_name= "unmerged_output.xlsx"
    issues = {'merge_empty_cells': []}

    #processing for a single file
    status=processing_excel(input_file_name,sheet_name,output_file_name,issues)

    #checking final status
    if status:
        print(f"✅ Successfully processed '{input_file_name}' -> '{output_file_name}'")
        print("Issues:",issues)
    else:
        print(f"❌ Failed to process '{input_file_name}'")