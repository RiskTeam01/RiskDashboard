from app.automations.credit_worksheet.fields import FIELD_DEFINITIONS
from app.config import SHEET_NAME


def get_status_metadata() -> dict:
    return {
        "field_count": len(FIELD_DEFINITIONS),
        "expected_sheet": SHEET_NAME,
    }
