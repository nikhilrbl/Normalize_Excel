# Project: Normalize Excel
# Description: Script to process Excel sheets by unmerging merged cells,
#              filling values, applying formatting, and saving the result.
# Author: EHCIKNA

from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
# from datetime import datetime
import os

# ================================
# Color Definitions for Cell Fills
# ================================
# These colors can be applied to highlight issues or mark cells.
red_fill     = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")   # Red
green_fill   = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")   # Green
blue_fill    = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")   # Blue
yellow_fill  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")   # Yellow
pink_fill    = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")   # Pink

# Extra optional colors
orange_fill  = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")   # Orange
purple_fill  = PatternFill(start_color="800080", end_color="800080", fill_type="solid")   # Purple
gray_fill    = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")   # Gray


def formatting(ws):
    """
    Apply center alignment to all non-empty cells in the given worksheet.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The worksheet to format.

    Returns
    -------
    bool
        True if formatting applied successfully, False if an error occurred.
    """
    try:
        # Define center alignment style
        center_align = Alignment(horizontal='center', vertical='center', text_rotation=0)

        # Apply alignment to all non-empty cells
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cell.alignment = center_align
        return True
    except Exception as e:
        print(f">> Function:{formatting.__name__}, Error processing file: {str(e)}")
        return False


def unmerge_fill(ws, issues):
    """
    Unmerge all merged cells in a worksheet and fill resulting cells.

    - If the merged cell had a value, propagate it to all unmerged cells.
    - If the merged cell was empty, highlight those cells with red fill.

    Parameters
    ----------
    ws : openpyxl.worksheet.worksheet.Worksheet
        The worksheet to process.
    issues : dict
        A dictionary to record problematic cases (e.g., empty merged cells).

    Returns
    -------
    bool
        True if unmerge and fill completed successfully, False otherwise.
    """
    try:
        # Iterate over all merged cell ranges
        for merge_range in list(ws.merged_cells.ranges):
            # Extract the top-left cell value of the merged range
            cell_value = ws[merge_range.coord.split(":")[0]].value
            if cell_value is None:
                issues['merge_empty_cells'].append(merge_range)

            # Unmerge the cell range
            ws.unmerge_cells(str(merge_range))

            # Fill each cell in the range with either value or red highlight
            for row in ws[merge_range.coord]:
                for cell in row:
                    if cell_value is None:
                        cell.fill = red_fill  # Highlight empty merged cells
                    else:
                        cell.value = cell_value  # Copy merged value
        return True
    except Exception as e:
        print(f">> Function:{unmerge_fill.__name__}, Error processing file: {str(e)}")
        return False


def processing_excel(*file_info):
    """
    Process an Excel file:
    - Unmerge all merged cells.
    - Fill unmerged cells with original merged values.
    - Apply center alignment formatting.
    - Save the result to a new file.

    Parameters
    ----------
    *file_info : tuple
        (input_file_name, sheet_name, output_file_name, issues)

    Returns
    -------
    bool
        True if processing succeeded, False otherwise.
    """
    input_file_name = file_info[0]
    sheet_name = file_info[1]
    output_file_name = file_info[2]
    issues = file_info[3]

    try:
        # Check input file existence
        if not os.path.exists(input_file_name):
            print(f">> '{input_file_name}' does not exist")
            return False

        # Load workbook
        wb = load_workbook(input_file_name, read_only=False)

        # Validate sheet name
        if sheet_name not in wb.sheetnames:
            print(f">> '{sheet_name}' does not exist. Available sheet names are \"{wb.sheetnames}\"")
            return False
        ws = wb[sheet_name]

        # Step 1: Unmerge merged cells
        unmerge_status = unmerge_fill(ws, issues)
        if not unmerge_status:
            print(f">> Unmerge Failed")
            return False

        # Step 2: Apply formatting
        align_status = formatting(ws)
        if not align_status:
            print(f">> Alignment Failed")
            return False

        # Step 3: Save processed workbook
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
    # ======================
    # Script Configuration
    # ======================
    input_file_name = "input.xlsx"                     # Input Excel file
    sheet_name = "Node Version Planner"                # Target sheet
    # file_datetime=datetime.now().strftime("%d-%m-%Y_%H-%M-%S") # Example timestamp format
    # output_file_name= "unmerged_output_" + str(file_datetime) + ".xlsx"
    output_file_name = "unmerged_output.xlsx"          # Output Excel file
    issues = {'merge_empty_cells': []}                 # Track issues like empty merged cells

    # Run processing for a single file
    status = processing_excel(input_file_name, sheet_name, output_file_name, issues)

    # Final status check
    if status:
        print(f"✅ Successfully processed '{input_file_name}' -> '{output_file_name}'")
        print("Issues:", issues)
    else:
        print(f"❌ Failed to process '{input_file_name}'")