from form import Form
import pandas


class SE361_Form(Form):
    """Form implementation for SE361 sprint reviews"""

    # Indices of columns from the raw excel sheet (zero-indexed)
    start_time_idx = 1
    section_idx = 5
    evaluator_type_idx = 6
    sprint_num_idx = 7
    team_name_idx = 8
    team_leader_idx = 9
    strength_comment_idx = 11
    improvement_comment_idx = 12
    extension_comment_idx = 13
    score_idx = 14
    explanation_comment_idx = 15

    # Indices reordered for output files
    column_order = [
        section_idx,
        evaluator_type_idx,
        sprint_num_idx,
        team_name_idx,
        team_leader_idx,
        score_idx,
        strength_comment_idx,
        improvement_comment_idx,
        extension_comment_idx,
        explanation_comment_idx,
    ]

    # Renamed columns for consistency
    new_column_names = [
        "Section",
        "Evaluator",
        "Sprint_Number",
        "Team_Name",
        "Team_Leader",
        "Scores",
        "Strengths",
        "Improvements",
        "Extensions",
        "Explanations",
    ]

    def __init__(self, input_file_path):
        super().__init__(input_file_path)
        self._read_path()

    def print_forms(self):
        self._fix_date_parsing()

        self._reorder_columns()

        for section, section_df in self.data_frame.groupby("Section", sort=False):
            # Loop through each individual team's review
            for team_name, team_df in section_df.groupby("Team_Name", sort=False):
                reordered_df = self._reorder_evaluator_and_average_rows(team_df)

                # Parse data frame for output filename data
                team_name = team_name.replace("Team", "").strip()
                section = section.strip()
                file_suffix = f"{team_name}_{section}"

                out_file = self.input_file_path.with_name(
                    f"{self.input_file_path.stem}_Team_{file_suffix}.csv"
                )

                print(f"Writing output file {out_file}")

                reordered_df.to_csv(out_file, index=False)

    def filter_by_date(self, date):
        """
        Filters data frame rows to all after input date
        """
        start_time = pandas.Timestamp(date)

        # Filter data by sprint number and date
        self.data_frame = self.data_frame[
            self.data_frame.iloc[:, SE361_Form.start_time_idx] >= start_time
        ]

    def filter_by_sprint(self, sprint_number):
        """
        Filters data frame rows by input sprint number
        """
        sprint_string = "Sprint " + sprint_number

        # Filter data by sprint number and date
        self.data_frame = self.data_frame[
            (self.data_frame.iloc[:, SE361_Form.sprint_num_idx] == sprint_string)
        ]

    def _reorder_evaluator_and_average_rows(self, data_frame):
        """
        Combines peer, peer average score, empty,
        stakeholder, and stakeholder average score rows
        """
        peer_rows, peer_avg_row = SE361_Form._get_peer_rows(data_frame)
        stakeholder_rows, stakeholder_avg_row = SE361_Form._get_stakeholder_rows(
            data_frame
        )

        # Create two blank rows matching dataframe structure
        blank_row = self._get_blank_rows(1)

        # Return final ordered data frame
        return pandas.concat(
            [peer_rows, peer_avg_row, blank_row, stakeholder_rows, stakeholder_avg_row],
            ignore_index=True,
        )

    def _get_peer_rows(data_frame):
        df = data_frame
        peer_rows = df[df["Evaluator"] == "Peer"]
        peer_avg_row = SE361_Form._get_avg_row(peer_rows)
        return peer_rows, peer_avg_row

    def _get_stakeholder_rows(data_frame):
        df = data_frame
        stakeholder_rows = df[df["Evaluator"] == "Stakeholder"]
        stakeholder_avg_row = SE361_Form._get_avg_row(stakeholder_rows)
        return stakeholder_rows, stakeholder_avg_row

    def _get_blank_rows(self, row_count=1):
        df = self.data_frame
        blank_rows = pandas.DataFrame(
            [[None] * len(df.columns) * row_count], columns=df.columns
        )
        return blank_rows

    def _get_avg_row(rows):
        avg = SE361_Form._calculate_average_score(rows)
        avg_row = {col: "" for col in rows.columns}
        avg_row["Team_Leader"] = "AVERAGE SCORE"
        avg_row["Scores"] = avg
        avg_row = pandas.DataFrame([avg_row])
        return avg_row

    def _calculate_average_score(df):
        """Calculates average score of input data frame"""
        avg = df["Scores"].mean()
        return round(avg, 3)

    def _fix_date_parsing(self):
        # Parse dates correctly
        self.data_frame.iloc[:, SE361_Form.start_time_idx] = pandas.to_datetime(
            self.data_frame.iloc[:, SE361_Form.start_time_idx], errors="coerce"
        )

    def _read_path(self):
        # Read the input Excel file as a data frame object
        self.data_frame = pandas.read_excel(
            self.input_file_path, dtype=object, parse_dates=True
        )

    def _reorder_columns(self):
        # Renames and reorders the columns
        self.data_frame = self.data_frame.iloc[:, SE361_Form.column_order]
        self.data_frame.columns = SE361_Form.new_column_names
