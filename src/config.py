import pandas as pd
from pathlib import Path
# =========================================================
# FILE PATHS
# =========================================================
#Input File path #
file_path = Path("data/Staff Planning 2027.xlsx")

#Template File Paht#
template_path = Path("template/template_return.xlsx")

# =========================================================
# PLANNING SECTIONS
# =========================================================
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

# =========================================================
# BUSINESS RULES
# =========================================================
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