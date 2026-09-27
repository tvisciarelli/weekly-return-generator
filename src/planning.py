import pandas as pd
from pathlib import Path

"""Find the Date Header and return its location and the dates to its right. """
def find_dates(df: pd.DataFrame) -> tuple[int, int, list]:
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

""""Find the row position of each known section."""
def find_sections(df: pd.DataFrame,section_names:list[str]) -> dict[str,int]:
    sections ={}

    for section in section_names:
        locations = df.eq(section).stack()

        if not locations.any():
            continue

        row, _ = locations[locations].index[0]
        sections[section] = row
    
    return sections

""""Find the range of each known section."""
def find_section_ranges(sections, df):
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