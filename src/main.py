# IMPORTS
import pandas as pd

from config import (
    file_path,
    template_path,
    section_names,
    sections_to_process
)

from rules import (
    apply_rules_staff,
    apply_rules_equipment
)

from planning import (
    find_dates,
    find_sections,
    find_section_ranges
)

from processing import (
    process_staff,
    process_equipment,
)

from output import (
    create_return_file
)


# Function to get Week Start and Validation ##
def get_week_start():
    while True:
        date_input = input("Enter the week start date (YYYY-MM-DD): ")

        try:
            week_start = pd.Timestamp(date_input)
        except ValueError:
            print("Invalid date. Please use YYYY-MM-DD")
            continue
        if week_start.weekday() != 0:
            print("The date must be a Monday.")
            continue
        

        return week_start

### MAIN ###
def main():
    #Request start date
    week_start = get_week_start()

    week_dates = pd.date_range(
        start=week_start,
        periods=7,
        freq="D"
    )

    #Load df#
    df = pd.read_excel(
        file_path,
        sheet_name="General Planning",
        header=None
    )
    
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

    equipment_tables = [marine_week,project_week,land_week]
    
    #Create Output (Personnel and Plant Return) #
    create_return_file(
    staff_week=staff_week,
    equipment_tables=equipment_tables,
    week_dates=week_dates,
    template_path=template_path,
)


if __name__ == "__main__":
    main()
