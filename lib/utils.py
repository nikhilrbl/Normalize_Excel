# Project: Node Details Parser
# Module: Utils
# Description: Utility Functions to reduce repeated code
# Author: EIMACAH

import os
import logging
from datetime import datetime


# -------------------------------
# Utility Functions
# -------------------------------

def get_last_row_with_value(ws):
    """Find the last row containing any non-empty value."""
    for row in range(ws.max_row, 0, -1):
        if any(cell.value not in (None, "") for cell in ws[row]):
            logging.debug(f"Last row with value: {row}")
            return row
    logging.debug("No rows with values found")
    return 0


def get_last_col_with_value(ws):
    """Find the last column containing any non-empty value."""
    for col in range(ws.max_column, 0, -1):
        if any(
            ws.cell(row=row, column=col).value not in (None, "")
            for row in range(1, ws.max_row + 1)
        ):
            logging.debug(f"Last column with value: {col}")
            return col
    logging.debug("No columns with values found")
    return 0


def setup_logging(output_base_name, log_option):
    """
    Configure file logging if -l option is used.
    Log file will be saved in the 'log/' folder with timestamp.
    If log_option is False/None, logging will be completely disabled.
    """
    if not log_option:
        # Disable logging entirely if user didn't request it
        logging.disable(logging.CRITICAL)
        return

    # Base folder of the project (relative to this utils.py file)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    log_dir = os.path.join(base_dir, "log")

    # Ensure folders exist
    os.makedirs(log_dir, exist_ok=True)

    # Determine base name for log file
    if isinstance(log_option, str):
        # Remove any directory and extension, keep only the base name
        base = os.path.splitext(os.path.basename(log_option))[0]
    else:
        # Default log filename based on output_base_name
        base = f"{output_base_name}_log"

    # Add timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{base}_{timestamp}.log"

    # Full path in log folder
    log_file = os.path.join(log_dir, log_file_name)

    # Configure logging: file only, no console
    logging.basicConfig(
        filename=log_file,
        # level=logging.INFO,
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filemode='w',
        force=True  # Clears any previous handlers
    )

    logging.info(f">> Logging started: {log_file}")
    print(f">> Logging enabled. File: {log_file}")
