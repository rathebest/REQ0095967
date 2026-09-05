#!/bin/sh
# =============================================================================
#  REQ0095968 - สร้างไฟล์สำหรับ Import เข้า PowerBuilder 8.0
#
#  ไฟล์ต้นฉบับใน pb/ เก็บเป็น UTF-8 + LF เพื่อให้อ่านและ diff บน Git ได้สะดวก
#  แต่ PowerBuilder 8.0 บน Windows ภาษาไทยอ่านไฟล์เป็น ANSI (TIS-620 / CP874)
#  และคาดหวังบรรทัดแบบ CRLF สคริปต์นี้จึงแปลงให้ก่อนนำไป Import
#
#  ใช้งาน :  sh tools/make_import_files.sh  [output_dir]
#  ค่าเริ่มต้น output_dir = build/import
# =============================================================================
set -e

OUT="${1:-build/import}"
SRC="pb"

if [ ! -d "$SRC" ]; then
    echo "ไม่พบโฟลเดอร์ $SRC - กรุณารันจากรากของ repository" >&2
    exit 1
fi

mkdir -p "$OUT"
n=0

for f in "$SRC"/*.srd "$SRC"/*.srw; do
    [ -e "$f" ] || continue
    base=$(basename "$f")
    # UTF-8 -> CP874 (TIS-620 + ส่วนขยายของ Windows) แล้วเติม CR ท้ายบรรทัด
    iconv -f UTF-8 -t CP874 "$f" | sed 's/$/\r/' > "$OUT/$base"
    n=$((n + 1))
    echo "  $base"
done

echo ""
echo "แปลงเสร็จ $n ไฟล์ -> $OUT (TIS-620/CP874, CRLF)"
echo "นำเข้าใน PowerBuilder 8.0 : Library Painter > เลือก cle_report.pbl > Import"
