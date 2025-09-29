# Project: Node Details Parser
# Description: Normalize Excel sheets and optionally create JSON output.
# Author: EIMACAH

import os
from datetime import datetime
import argparse
from openpyxl import load_workbook
from lib.pre_process import processing_excel
from lib.post_process import create_hierarchical_json
from lib.utils import setup_logging

def main():
    # ======================
    # Argument Parsing
    # ======================
    parser = argparse.ArgumentParser(
        description="Convert Excel node details into a hierarchical JSON file."
    )

    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path to the input Excel file (e.g., node_details.xlsx)"
    )
    parser.add_argument(
        "--sheetname", "-x",
        required=True,
        help="Sheet name in the Excel file (e.g., 'Node Version Planner')"
    )
    parser.add_argument(
        "--output", "-o",
        required=False,
        help="Output Excel file path (optional, default: InputFileName_date_time.xlsx)"
    )
    parser.add_argument(
        "--start-version", "-s",
        required=False,
        default="",
        help="Optional: Start version for JSON creation (default: first column)"
    )
    parser.add_argument(
        "--end-version", "-e",
        required=False,
        default="",
        help="Optional: End version for JSON creation (default: last column)"
    )
    parser.add_argument(
        "--json", "-j",
        nargs="?",
        const=True,
        help="Optional: Create JSON file. Use -j for default filename or -j filename.json for custom name"
    )
    parser.add_argument(
        "--logging", "-l",
        nargs="?",
        const=True,
        help="Optional: Enable logging. Use -l for default log filename or -l filename.log for custom name"
    )

    args = parser.parse_args()

    # ======================
    # Prepare Output Directory
    # ======================
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)

    # ======================
    # Determine Excel output filename
    # ======================
    if args.output:
        # Use user-provided Excel output path as-is
        output_file_name = args.output
    else:
        # Default Excel output: input base + timestamp
        output_base = os.path.splitext(os.path.basename(args.input))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file_name = os.path.join(output_dir, f"{output_base}_{timestamp}.xlsx")

    # ======================
    # Setup Logging
    # ======================
    setup_logging(os.path.splitext(os.path.basename(output_file_name))[0], args.logging)

    # ======================
    # Initialize issues dictionary
    # ======================
    issues = {
        'merged_empty_cells': [],
        'empty_cell_in_enm_version_row2': [],
        'empty_cells_after_unmerge': [],
        'unusable_rows': [],
        'incomplete_rows': [],
        'node_header_rows': [],
        # 'removed_header_rows': []
    }
    skipped_rows_for_json = []

    # ======================
    # Process Excel
    # ======================
    status = processing_excel(args.input, args.sheetname, output_file_name, issues)

    if not status:
        print(f">> Failed to process '{args.input}'")
        return

    print(f">> Successfully processed '{args.input}' -> '{output_file_name}'\n")

    print(">> Issues found during processing:")
    for key, value in issues.items():
        print(f"--{key.upper()} : \n{value}\n")
    print("")

    # ======================
    # Create JSON (if requested)
    # ======================
    if args.json:
        if args.json is True:
            # Default JSON name: same base as Excel, no timestamp
            json_base = os.path.splitext(os.path.basename(output_file_name))[0]
            json_file_name = os.path.join(output_dir, f"{json_base}.json")
        else:
            # User provided JSON filename
            json_file_name = os.path.join(output_dir, args.json)

        json_status = create_hierarchical_json(
            load_workbook(output_file_name)[args.sheetname],
            json_file_name,
            skipped_rows_for_json,
            args.start_version,
            args.end_version
        )

        if json_status:
            print(f">> Output JSON created: '{json_file_name}'\n")
            if skipped_rows_for_json:
                print(">> Skipped rows for JSON creation:")
                print(skipped_rows_for_json)
        else:
            print(f">> Failed to create JSON file")

if __name__ == "__main__":
    main()
