import pandas as pd

from rules import (
    apply_rules_staff,
    apply_rules_equipment,
)

### HELPER FUNCTIONS ####
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

### STAFF FUNCTION ####
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

### EQUIPMENT FUNCTION ###
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