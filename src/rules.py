import pandas as pd

from config import STAFF_RULES, EQUIPMENT_RULES


"""Normalize a raw status value"""
def normalize_status(status):
    if pd.isna(status):
        return None
    
    return str(status).strip().upper()


"""Apply Staff-specific status rules."""
def apply_rules_staff(df):
    def get_staff_rule(status):
        status = normalize_status(status)
        return STAFF_RULES.get(status, (None, None))
    
    df[["Hours", "Remark"]] = df["Status"].apply(
        lambda status: pd.Series(get_staff_rule(status))
    )

    return df

"""Apply common equipment/plant status rules."""
def apply_rules_equipment(df):
    
    def get_equipment_rule(status):
        status = normalize_status(status)
        return EQUIPMENT_RULES.get(status)

    df["Processed_Status"] = df["Status"].apply(get_equipment_rule)

    return df