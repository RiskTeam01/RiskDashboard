# Each entry defines one Excel output cell and the FOCUS helper code(s) to extract from the PDF.
# expression: a single code or "CODE1+CODE2" for a summed field
# label:      human-readable description
# excel_cell: target cell in the output workbook (Sheet1)
FIELD_DEFINITIONS = [
    {"expression": "940",        "label": "Total Assets",                         "excel_cell": "B4"},
    {"expression": "1800",       "label": "Total Equity",                         "excel_cell": "B6"},
    {"expression": "750",        "label": "Cash and cash equivalents",             "excel_cell": "B8"},
    {"expression": "760",        "label": "Cash segregated under federal",         "excel_cell": "B10"},
    {"expression": "770",        "label": "Fails to Deliver",                      "excel_cell": "B12"},
    {"expression": "780",        "label": "Stocks Borrowed",                       "excel_cell": "B14"},
    {"expression": "800",        "label": "Clearing Org receivables",              "excel_cell": "B16"},
    {"expression": "810",        "label": "Others",                                "excel_cell": "B18"},
    {"expression": "820",        "label": "Customer Receivables",                  "excel_cell": "B20"},
    {"expression": "840",        "label": "Reverse Repos",                         "excel_cell": "B22"},
    {"expression": "292",        "label": "Trade Date Receivable",                 "excel_cell": "B24"},
    {"expression": "12019",      "label": "Marketable securities",                 "excel_cell": "B26"},
    {"expression": "740",        "label": "Non-allowable assets",                  "excel_cell": "B28"},
    {"expression": "890",        "label": "Secured Demand Notes",                  "excel_cell": "B32"},
    {"expression": "1760",       "label": "Total Liabilities",                     "excel_cell": "B34"},

    {"expression": "1490+1500",  "label": "Fails to Receive",                      "excel_cell": "D12"},
    {"expression": "1510+1520",  "label": "Stocks Loaned",                         "excel_cell": "D14"},
    {"expression": "1550+1560",  "label": "Clearing Org payables",                 "excel_cell": "D16"},
    {"expression": "1570",       "label": "Other",                                 "excel_cell": "D18"},
    {"expression": "1580+1590",  "label": "Customer payables",                     "excel_cell": "D20"},
    {"expression": "1480",       "label": "Repos",                                 "excel_cell": "D22"},
    {"expression": "1686",       "label": "Obligation to rtn Securities Collateral","excel_cell": "D26"},
    {"expression": "1480",       "label": "All other liabilities",                 "excel_cell": "D30"},
    {"expression": "1730",       "label": "Securities borrowings",                 "excel_cell": "D32"},
]
