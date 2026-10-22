"""Lightweight shared attendance categories and month names."""

MONTHS = (
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
)

BUDDHIST_YEAR_MIN = 1000
BUDDHIST_YEAR_MAX = 9999


def is_valid_buddhist_year(year):
    """Return whether a Buddhist Era year is exactly four ASCII digits."""
    value = str(year)
    return (
        value.isascii()
        and value.isdigit()
        and len(value) == 4
        and BUDDHIST_YEAR_MIN <= int(value) <= BUDDHIST_YEAR_MAX
    )

TEMPLATES = {
    "temporary": {
        "Cover": "ปกลูกจ้าง.docx",
        "Table": "บัญชีลงเวลาการปฏิบัติราชการของลูกจ้าง.docx",
    },
    "government": {
        "Cover": "ปกพนักงานราชการ.docx",
        "Table": "บัญชีลงเวลา-พนักงานราช.docx",
    },
    "guard": {
        "Cover": "ปกยาม.docx",
        "Table": "บัญชีลงเวลาปฏิบัติหน้าที่ยามรักษาการณ์ป.doc",
    },
}
