import shutil
from copy import copy

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment

from pathlib import Path

### --------------------------------------
# ADD STYLE TO TEMPLATE
# ---------------------------------------
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


###---------------------------------------
# WRITE PERSONNEL
###---------------------------------------
"""Function to write personnel table in template"""

def write_personnel(ws,staff_week,week_dates):
    week_start = week_dates[0]
    """Add week"""
    ws.cell(
        row=3,
        column=14
    ).value = week_start.isocalendar().week
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


###---------------------------------------
# WRITE EQUIPMENT
###---------------------------------------
def write_equpment(ws, equipment_tables, week_dates):
    week_start = week_dates[0]
    """Add week"""
    ws.cell(
        row=3,
        column=16
    ).value = week_start.isocalendar().week
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
    """Add Tables With Equipment"""
    equipment_start_row = 9
    equipment_start_column = 2
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


###---------------------------------------
# WRITE FINAL RETURN
###---------------------------------------
"""Function to create return file"""
def create_return_file(staff_week,equipment_tables,week_dates,template_path):
    week_start = week_dates[0]
    # 1. Determine output filename
    output_filename = f"weekly_return_{week_start.isocalendar().year}_Wk{week_start.isocalendar().week}.xlsx"
    output_path = Path("output")/output_filename
    
    # 2. Copy template
    shutil.copy2(template_path, output_path)

    # 3. Open copied workbook
    wb = load_workbook(template_path)
    
    # 4. Get worksheets
    ws_personnel = wb["Personnel"]
    ws_equipment = wb["Plant"]

    # 5. Write personnel
    write_personnel(ws_personnel,staff_week,week_dates)
    
    # 6. Write equipment
    write_equpment(ws_equipment,equipment_tables,week_dates)

    # 7. Save workbook
    wb.save(output_path)


    


