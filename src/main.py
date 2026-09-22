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


def find_dates(df: pd.DataFrame) -> tuple[int, int, list]:
    """Find the Date Header and return its location and the dates to its right. """
    locations = df.eq("Date").stack()

    if locations.empty:
        raise ValueError("Could not find 'Date' in the planning sheet.")
    
    row, column = locations[locations].index[0]

    date_values = df.iloc[row, column+1:]

    dates = pd.to_datetime(
        date_values,
        errors='coerce'
    ).dropna().tolist()

    return row, column, dates


def find_sections(df: pd.DataFrame,section_names:list[str]) -> dict[str,int]:
    """"Find the row position of each known section."""
    sections ={}

    for section in section_names:
        locations = df.eq(section).stack()

        if not locations.any():
            continue

        row, _ = locations[locations].index[0]
        sections[section] = row
    
    return sections

def find_section_ranges(sections, df):
    """"Find the range of each known section."""
    section_ranges = {}

    sections = list(sections.items())
    sections = sorted(sections, key=lambda x:x[1])

    for i, (name,start) in enumerate(sections):
        if i < len(sections) - 1:
            end = sections[i+1][1]
        else:
            end = df.shape[0]
            
        section_ranges[name]=(start,end)
        
    return section_ranges


def main():
    file_path = Path("data/20.4045 Staff Planning 2027.xlsx")

    df = load_planning_sheet(file_path)

    date_row, date_column, dates = find_dates(df)

    section_names = ["Holidays",
                     "Main Activities Planning",
                     "Staff Planning","Accommodation Planning",
                     "Marine Equipment Planning",
                     "Project Equipment",
                     "Land Based Equipment Planning",
                     "Activities Planning"]
    
    sections_to_process = ["Staff Planning",
                           "Marine Equipment Planning",
                           "Project Equipment",
                           "Land Based Equipment Planning"]
    

    
    sections = find_sections(df, section_names)

    section_ranges = find_section_ranges(sections,df)

    #Filter the sections that we want to use:
    selected_sections = {
        name: range
        for name, range in section_ranges.items() if name in sections_to_process
    }

    print(f"Date header found at row {date_row}, column {date_column}")
    print(f"Sections: {sections}")
    print()
    print(f"Section Ranges: {section_ranges}")
    print()
    print(f"Selected Sections: {selected_sections}")

if __name__ == "__main__":
    main()
