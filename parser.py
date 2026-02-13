import argparse
import pandas as pd
from pathlib import Path

# Input date to filter forward from using format yyyy-mm-dd
DATE_FILTER = "2026-01-01"

# Input sprint number to filter submissions by
SPRINT_NUM_FILTER = "0"

# Indices for all columns needed in output file
start_time_idx = 1
sprint_num_idx = 6
team_name_idx = 7
team_leader_idx = 8
score_idx = 13
strength_comment_idx = 10
improvement_comment_idx = 11
extension_comment_idx = 12
explanation_comment_idx = 14

needed_columns = [sprint_num_idx, team_name_idx, team_leader_idx, score_idx,
                  strength_comment_idx, improvement_comment_idx,
                  extension_comment_idx, explanation_comment_idx]

def main():
    input_path = get_input_path()

    # Read the input Excel file as a data frame object
    df = pd.read_excel(input_path, dtype=object, parse_dates=True)

    # Parse dates correctly
    df.iloc[:, start_time_idx] = pd.to_datetime(df.iloc[:,start_time_idx], errors="coerce")

    filtered_df = filter_data_by_sprint_number_and_date(df, SPRINT_NUM_FILTER, DATE_FILTER)

    # Extract needed columns and name them
    subset = filtered_df.iloc[:, needed_columns]
    subset.columns = ["Sprint_Number","Team_Name","Team_Leader", "Scores","Strengths","Improvements","Extensions","Explanations"]

    # Loop through each individual team's review
    for team_name, team_df in subset.groupby("Team_Name", sort=False):
        # Calculate the average score given
        avg = team_df["Scores"].mean()
        avg = round(avg, 3)
        suffix = team_name.replace("Team", "").strip()

        # Build average score row
        avg_row = {col: "" for col in team_df.columns}
        avg_row["Team_Leader"] = "AVERAGE SCORE"
        avg_row["Scores"] = avg

        # Concatenate the average score row to this team's data frame
        team_df = pd.concat([team_df, pd.DataFrame([avg_row])], ignore_index=True)

        out_file = input_path.with_name(
            f"{input_path.stem}_Team_{suffix}.csv"
        )

        team_df.to_csv(out_file, index=False)


def filter_data_by_sprint_number_and_date(df, sprint_number, date):
    """
    Returns a data file with rows only equal to input sprint number
    and after input date
    """
    sprint_string = "Sprint " + sprint_number
    start_time = pd.Timestamp(date)

    # Filter data by sprint number and date
    filtered_df = df[
        (df.iloc[:, sprint_num_idx] == sprint_string) &
        (df.iloc[:, start_time_idx] >= start_time)
    ]

    return filtered_df

def get_input_path():
    """ RETURN INPUT FILEPATH FROM COMMAND LINE  """
    parser = argparse.ArgumentParser(
        description="Split an Excel file into separate files based on the Team column."
    )
    parser.add_argument(
        "input_file",
        help="Path to the input Excel file"
    )

    args = parser.parse_args()
    input_path = Path(args.input_file)

    if not input_path.exists():
        raise FileNotFoundError(f"File not found: {input_path}")

    return input_path


if __name__ == "__main__":
    main()
