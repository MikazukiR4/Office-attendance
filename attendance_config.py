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
