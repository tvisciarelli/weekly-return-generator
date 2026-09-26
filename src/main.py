from pathlib import Path
import pandas as pd

###RULES AND INPUT###
STAFF_RULES = {
    "A": (None, None),
    "WFH": (8, "Working from home"),
    "1": (12, None),
    "T": (12, None),
    "2": (6, None)
}

EQUIPMENT_RULES = {
    "1": "x",
    "M": "m",
    "S": "s"
}

week_start = pd.Timestamp("2027-06-14")
week_dates = pd.date_range(start = week_start,
                           periods=7,
                           freq="D")



###GENERAL FUNCTIONS###
def load_planning_sheet(file_path: Path) -> pd.DataFrame:
    """"Load the General Planning sheet"""
    return pd.read_excel(
        file_path,
        sheet_name="General Planning",
        header=None
    )

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

def extract_section(df, section_range):
    extract_df = df.iloc[section_range[0]+1:section_range[1],:].copy()
    extract_df = extract_df.reset_index(drop=True)

    return extract_df

def clean_section(df,titles):
    df = df.dropna(axis=0,subset=[1,2,3])
    df = df.drop([0,4],axis=1).reset_index(drop=True)
    df = df.copy()
    df.insert(0, 'Source_Order', range(len(df)))
    assert len(titles) == len(df.columns), (
        f"Number of titles ({len(titles)}) does not match "
        f"number of columns ({len(df.columns)})"
    )
        
    df.columns = titles
    
    df = df.melt(
        id_vars=titles[0:4],
        var_name="Date",
        value_name='Status'
    )

    return df

def normalize_status(status):
    """Normalize a raw status value."""
    if pd.isna(status):
        return None
    
    return str(status).strip().upper()
    
### STAFF FUNCTIONS ####
def apply_rules_staff(df):
    """Apply Staff-specific status rules."""
    def get_staff_rule(status):
        status = normalize_status(status)
        return STAFF_RULES.get(status, (None, None))
    
    df[["Hours", "Remark"]] = df["Status"].apply(
        lambda status: pd.Series(get_staff_rule(status))
    )

    return df

def process_staff(df, section_range, titles, week_dates):
    df_staff = extract_section(df,section_range)
    df_staff = clean_section(df_staff, titles)
    df_staff = apply_rules_staff(df_staff)

    df_staff = df_staff[df_staff["Date"].isin(week_dates)]

    active = (
        df_staff
        .groupby(["Source_Order","ID_1","ID_2","ID_3"])["Hours"]
        .transform(lambda x: x.notna().any())
    )

    df_staff = df_staff[active]

    remarks = (
            df_staff[df_staff["Remark"].notna()]
            .groupby(["Source_Order","ID_1", "ID_2", "ID_3"])["Remark"]
            .agg(lambda x: ", ".join(x.unique()))
            .reset_index()
    )

    df_staff['Hours'] = df_staff['Hours'].astype('Int64')

    df_staff = (df_staff
                .pivot(
                    index=['Source_Order','ID_1','ID_2','ID_3'],
                    columns='Date',
                    values='Hours')
                .reindex(
                    columns=week_dates)
                .reset_index()
                .sort_values("Source_Order"))

    df_staff["Total Hours"] = df_staff[week_dates].sum(axis=1).astype('Int64')


    df_staff = df_staff.merge(
                remarks,
                on=['Source_Order',"ID_1", "ID_2", "ID_3"],
                how="left",
                ).drop(columns='Source_Order')
    
    return df_staff


### EQUIPMENT FUNCTIONS ###
def apply_rules_equipment(df):
    """Apply common equipment/plant status rules."""

    def get_equipment_rule(status):
        status = normalize_status(status)
        return EQUIPMENT_RULES.get(status)

    df["Processed_Status"] = df["Status"].apply(get_equipment_rule)

    return df

def process_equipment(df, section_range, titles, week_dates):
    df_equipment = extract_section(df,section_range)
    df_equipment = clean_section(df_equipment, titles)
    df_equipment = apply_rules_equipment(df_equipment)

    df_equipment = df_equipment[df_equipment["Date"].isin(week_dates)]

    active = (
        df_equipment
        .groupby(["Source_Order","ID_1","ID_2","ID_3"])["Processed_Status"]
        .transform(lambda x: x.notna().any())
    )

    df_equipment = df_equipment[active]

    df_equipment = (df_equipment
                .pivot(
                    index=['Source_Order','ID_1','ID_2','ID_3'],
                    columns='Date',
                    values="Processed_Status")
                .reindex(
                    columns=week_dates)
                .reset_index()
                .sort_values("Source_Order")
                .drop(columns='Source_Order'))
    
    df_equipment["Standby Days"] = (
        df_equipment[week_dates].eq("s").sum(axis=1)
    )

    df_equipment["Mob/Demob"] = (
        df_equipment[week_dates].eq("m").sum(axis=1)
    )

    df_equipment["Working Days"] = (
        df_equipment[week_dates].eq("x").sum(axis=1)
    )
    
    return df_equipment


### MAIN ###
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
        name: section_range
        for name, section_range in section_ranges.items() if name in sections_to_process
    }

    titles = ['Source_Order','ID_1', 'ID_2', 'ID_3'] + dates

    # Process Staff #
    staff_week = process_staff(
        df,
        selected_sections['Staff Planning'],
        titles,
        week_dates
    )
    
    # Process Marine Equipment #
    marine_week = process_equipment(
        df,
        selected_sections["Marine Equipment Planning"],
        titles,
        week_dates
    )

    # Process Project Equipment #
    project_week = process_equipment(
        df,
        selected_sections['Project Equipment'],
        titles,
        week_dates
    )

    #Process Land Based Equipment #
    land_week = process_equipment(
        df,
        selected_sections['Land Based Equipment Planning'],
        titles,
        week_dates
    )

    print(land_week.head())
    #print(marine_week.columns)

if __name__ == "__main__":
    main()
