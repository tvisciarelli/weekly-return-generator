from pathlib import Path
import pandas as pd

def load_planning_sheet(file_path: Path) -> pd.DataFrame:
    """"Load the General Planning sheet"""
    return pd.read_excel(
        file_path,
        sheet_name="General Planning",
        header=None
    )

def find_date_header(df: pd.DataFrame) -> tuple[int, int]:
    """Find the 'Date' cell and return its row and column."""
    locations = df.eq("Date").stack()

    row, column = locations[locations].index[0]
    return row, column


def main():
    file_path = Path("data/20.4045 Staff Planning 2027.xlsx")

    df = load_planning_sheet(file_path)

    date_row, date_column = find_date_header(df)

    print(f"Date header found at row {date_row}, column {date_column}")

if __name__ == "__main__":
    main()
