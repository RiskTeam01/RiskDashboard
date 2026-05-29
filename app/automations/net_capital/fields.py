# Maps row number in the Net Capital sheet to the FOCUS helper code.
# These were read directly from the template ec28fc95-Net_Capital.xlsx.
NET_CAPITAL_ROW_MAP: dict[int, str] = {
    7:  "3500",
    10: "3520",
    12: "3530",
    15: "3540",
    16: "3550",
    17: "3560",
    18: "12051",
    19: "12052",
    20: "3570",
    21: "3580",
    22: "3590",
    23: "3600",
    24: "3610",
    25: "3615",
    26: "3620",
    27: "3630",
    28: "3640",
    31: "3660",
    32: "3670",
    33: "3680",
    34: "3690",
    35: "3700",
    36: "3710",
    37: "3720",
    38: "3730",
    39: "3732",
    40: "12028",
    41: "3734",
    42: "3650",
    43: "3736",
    44: "12053",
    45: "12054",
    46: "3740",
    47: "3750",
    50: "3760",
    53: "3910",
    59: "3870",
    60: "3880",
}

# Column letters C=Jan through N=Dec (0-indexed offset from C)
MONTH_COLUMNS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# Required quarter months (1-indexed): March=3, June=6, September=9, December=12
REQUIRED_MONTHS = {3, 6, 9, 12}

# Row 5 = period end date (code 25), Row 1 = company name (code 13)
PERIOD_END_ROW = 5
COMPANY_NAME_ROW = 1
