try:
    import argparse
    from pathlib import Path
    import pandas
    import openpyxl
except ModuleNotFoundError as e:
    raise SystemExit(
        f"\nMissing dependency: {e.name}\n"
        "Please install required packages first:\n"
        "    pip install -r requirements.txt\n"
    )

from se361_form import SE361_Form

# -------------------- FILTER OPTIONS --------------------

# Input date to filter forward from (date format: yyyy-mm-dd)
DATE_FILTER = "2026-02-24"

# Input sprint number to filter submissions by
SPRINT_NUM_FILTER = "2"

# ---------------------- END OPTIONS ---------------------


def main():
    input_path = get_input_path()

    # Read the input Excel file as a data frame object
    form = SE361_Form(input_path)

    # Filter the form data to reduce parsing time
    form.filter_by_date(DATE_FILTER)
    form.filter_by_sprint(SPRINT_NUM_FILTER)

    # Print the forms to output files in this project's directory
    form.print_forms()


def get_input_path():
    """Returns input file path from the command line arguments"""
    parser = argparse.ArgumentParser(
        description="Split an Excel file into separate files based on the Team column."
    )
    parser.add_argument("input_file", help="Path to the input Excel file")

    args = parser.parse_args()
    input_path = Path(args.input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")

    return input_path


if __name__ == "__main__":
    main()
