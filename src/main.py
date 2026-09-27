from pathlib import Path
import pandas as pd
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from copy import copy

###INITIAL REPORTING DATE###
week_start = pd.Timestamp("2027-06-24")
week_dates = pd.date_range(start = week_start,
                           periods=7,
                           freq="D")

iso = week_start.isocalendar()

#Input File path #
file_path = Path("data/Staff Planning 2027.xlsx")

#Template File Paht#
template_path = Path("data/template_return.xlsx")

#Output file#
output_filename = f"weekly_return_{iso.year}_Wk{iso.week}.xlsx"
output_path = Path("output")/output_filename
shutil.copy2(template_path, output_path)



### SECTIONS IN THE FILE AND SECTIONS TO PROCESS ###
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

###GENERAL FUNCTIONS###
"""Loads Planning File"""
def load_planning_sheet(file_path: Path) -> pd.DataFrame:
    """"Load the General Planning sheet"""
    return pd.read_excel(
        file_path,
        sheet_name="General Planning",
        header=None
    )

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

"""Extracts each of the sections based on the range"""
def extract_section(df, section_range):
    extract_df = df.iloc[section_range[0]+1:section_range[1],:].copy()
    extract_df = extract_df.reset_index(drop=True)

    return extract_df

"""General Cleanning of the Section"""
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

"""Normalize a raw status value"""
def normalize_status(status):
    if pd.isna(status):
        return None
    
    return str(status).strip().upper()

"""Add Style to Final Personnel Tables"""
def format_data_area(ws,start_row,end_row,start_col,end_col,remarks_column = 13, remarks_merge_column = 14):
    double_side = Side(style="double")
    dotted_side = Side(style="dotted")

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):

            cell = ws.cell(row=row, column=col)

            # -----------------------------------
            # Alignment
            # -----------------------------------

            cell.alignment = copy(
                cell.alignment
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            # -----------------------------------
            # Borders
            # -----------------------------------

            left = (
                double_side
                if col == start_col
                else cell.border.left
            )

            right = (
                double_side
                if col == end_col
                else cell.border.right
            )

            top = (
                double_side
                if row == start_row
                else dotted_side
            )

            bottom = (
                double_side
                if row == end_row
                else dotted_side
            )

            cell.border = Border(
                left=left,
                right=right,
                top=top,
                bottom=bottom,
            )

    # -----------------------------------
    # Merge Remarks columns
    # -----------------------------------

    for row in range(start_row, end_row + 1):
        ws.merge_cells(
            start_row=row,
            start_column=remarks_column,
            end_row=row,
            end_column=remarks_merge_column
        )


    
### STAFF FUNCTIONS ####
"""Apply Staff-specific status rules."""
def apply_rules_staff(df):
    def get_staff_rule(status):
        status = normalize_status(status)
        return STAFF_RULES.get(status, (None, None))
    
    df[["Hours", "Remark"]] = df["Status"].apply(
        lambda status: pd.Series(get_staff_rule(status))
    )

    return df

"""Process Staff Section to get final output"""
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

"""Function to write personnel table in template"""

def write_personnel(ws,staff_week,week_dates):
    """Add week"""
    ws.cell(
        row=3,
        column=14
    ).value = iso.week
    """Add Date From:"""
    ws.cell(
        row=4,
        column=14
    ).value = week_start
    """Add Date Until"""
    ws.cell(
        row=5,
        column=14
    ).value = week_dates[-1]
    """Add Reporting Dates"""
    for i, date in enumerate(week_dates):
        ws.cell(
            row=7,
            column=5+i
        ).value = date
    """Add Table With Personnel"""
    personnel_start_row = 9
    personnel_start_column = 2
    last_personnel_row = personnel_start_row + len(staff_week)-1
    last_personnel_column = 14

    for i, (_, person) in enumerate(staff_week.iterrows()):
        row = personnel_start_row + i
        ws.cell(row=row, column=2).value = person["ID_3"] #Company
        ws.cell(row=row, column=3).value = person["ID_2"] #Name
        ws.cell(row=row, column=4).value = person["ID_1"] #Role
        ws.cell(row=row, column=12).value = person["Total Hours"] #Total Hours
        ws.cell(row=row, column=13).value = person["Remark"] #Remark

        for j, date in enumerate(week_dates): #Daily Hours
            value = person[date]
            if pd.isna(value):
                value = None

            ws.cell(
                row=row,
                column=5 + j
            ).value = value

 # Apply format #
    format_data_area(
        ws,
        personnel_start_row,
        last_personnel_row,
        personnel_start_column,
        last_personnel_column,
    )




### EQUIPMENT FUNCTIONS ###
"""Apply common equipment/plant status rules."""
def apply_rules_equipment(df):
    
    def get_equipment_rule(status):
        status = normalize_status(status)
        return EQUIPMENT_RULES.get(status)

    df["Processed_Status"] = df["Status"].apply(get_equipment_rule)

    return df

"""Process equipment/plant sections"""
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

def write_equpment(ws, equipment_tables, week_dates):
    """Add week"""
    ws.cell(
        row=3,
        column=16
    ).value = iso.week
    """Add Date From:"""
    ws.cell(
        row=4,
        column=16
    ).value = week_start
    """Add Date Until"""
    ws.cell(
        row=5,
        column=16
    ).value = week_dates[-1]
    """Add Reporting Dates"""
    for i, date in enumerate(week_dates):
        ws.cell(
            row=7,
            column=5+i
        ).value = date
    """Add Table With Personnel"""
    equipment_start_row = 9
    equipment_start_column = 2
    #last_equipment_row = equipment_start_row + len(staff_week)-1
    last_equipment_column = 16
    empty_row = 1
    current_row = equipment_start_row

    for i, equipment_df in enumerate(equipment_tables):
        table_start_row = current_row
        # Write this table
        for _, equipment in equipment_df.iterrows():
            ws.cell(row=current_row, column=2).value = equipment["ID_3"] #Company
            ws.cell(row=current_row, column=3).value = equipment["ID_1"] #Equipment
            ws.cell(row=current_row, column=4).value = equipment["ID_2"] #Type
            ws.cell(row=current_row, column=12).value = equipment["Mob/Demob"] #Mab/Demob Days
            ws.cell(row=current_row, column=13).value = equipment["Standby Days"] #Standby Days
            ws.cell(row=current_row, column=14).value = equipment["Working Days"] #Working Days

            for j, date in enumerate(week_dates): #Daily Hours
                value = equipment[date]
                if pd.isna(value):
                    value = None

                ws.cell(
                    row=current_row,
                    column=5 + j
                ).value = value

            current_row += 1
        
        table_end_row = current_row - 1
        # Format this table #
        format_data_area(ws, table_start_row,table_end_row, equipment_start_column, last_equipment_column,remarks_column = 15, remarks_merge_column = 16)
        
        # Leave one blank eow before the next table
        if i < len(equipment_tables) - 1:
            current_row +=1

### MAIN ###
def main():    
    #Load df#
    df = load_planning_sheet(file_path)
    
    #Extract dates#
    date_row, date_column, dates = find_dates(df)
    
    #Find and extract section ranges#    
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

    #Write Personnel to Template#
    wb = load_workbook(template_path)
    ws_personnel = wb["Personnel"]
    ws_equipment = wb["Plant"]

    write_personnel(ws_personnel,staff_week, week_dates)
    write_equpment(ws_equipment,[marine_week,project_week,land_week],week_dates)

    wb.save(output_path)

if __name__ == "__main__":
    main()
