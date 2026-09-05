# ตารางสอบกลับความต้องการ — REQ0095968

สอบกลับทุกข้อกำหนดในเอกสาร CR REQ0095968 v1.2 (17/06/2569) ไปยังอ็อบเจกต์ที่ส่งมอบ

## ส่วนที่ 1 — เมนูและพารามิเตอร์

| CR | ข้อกำหนด | อ็อบเจกต์ที่รองรับ |
| --- | --- | --- |
| 1 | เพิ่มเมนูเรียกใช้งาน (สาขา ข้อ 53 / สนญ. ข้อ 38) | ผูกเมนูเข้ากับ `w_acvun7123` — ดู README ข้อ 3.3 |
| 2 | พารามิเตอร์ **ชื่อผู้ตัดชำระ** — Dropdown หรือเลือกทั้งหมด | `d_acvun7123_cond.operator_login` (DDDW `dddw_acvun7123_oper`) + `flag_all_oper` (Checkbox) · `proc_acvun7123_operator_list` |
| 2 | พารามิเตอร์ **สาขา** — ตาม Login แก้ไขไม่ได้ | `d_acvun7123_cond.br_code` / `br_name` (`protect="1"`, พื้นหลังเทา) · ตั้งค่าใน `w_acvun7123` event `open` จาก `gs_branch` |
| 2 | พารามิเตอร์ **เลขใบนำส่งฯ** From–To รูปแบบ `00000/00` | `debit_no_from` / `debit_no_to` (EditMask `#####/##`) · แปลงคีย์ด้วย `fn_acvun_debitno_key` |
| 2 | พารามิเตอร์ **วันที่ใบนำส่งฯ** From–To รูปแบบ `DD/MM/YYYY` พ.ศ. | `receipt_from` / `receipt_to` (EditMask `##/##/####`) · แปลงด้วย `wf_thai_to_datetime()` |
| 2 | Submit แล้วแสดง List ให้ผู้ใช้เลือกรายการ | `wf_search()` → `proc_acvun7123_search` → `d_acvun7123_list` |
| 2 | เลือกได้ตั้งแต่ 1 รายการขึ้นไปก่อน Save | `d_acvun7123_list.sel_flag` (Checkbox) · `wf_selected_rows()` · ปุ่ม เลือกทั้งหมด / ยกเลิกทั้งหมด (`wf_set_all()`) |

## ส่วนที่ 2 — รูปแบบไฟล์และการเรียงลำดับ

| CR | ข้อกำหนด | อ็อบเจกต์ที่รองรับ |
| --- | --- | --- |
| 3 | PDF A4 แนวนอน | `d_acvun7123_rpt` — `print.orientation = 1`, `print.paper.size = 9` |
| 3 | PDF สร้าง 1 ไฟล์ ต่อ 1 ใบนำส่งฯ (แยกไฟล์) | `wf_print_pdf()` — วนลูปรายการที่เลือก ตั้ง `Print.FileName` ใหม่ทุกรอบ |
| 3 | Excel สร้างเป็นไฟล์เดียว | `wf_export_excel()` → `wf_append_xls()` สะสมทุกใบลง `ids_xls` แล้ว `SaveAs(Excel5!)` ครั้งเดียว |
| 3 | หัวรายงานกึ่งกลางเมื่อพิมพ์ PDF | `d_acvun7123_rpt` band `header` ทุก object `alignment="2"` |
| 3 / เงื่อนไข 6 | เรียงเลขใบนำส่งฯ น้อย→มาก, Policy → Endorse, %ค่าใช้จ่าย น้อย→มาก | `proc_acvun7123_search` `ORDER BY debit_no` · `proc_acvun7123_summary` `ORDER BY rec_type, comm_rate, product_group` · `SetSort()` ใน `wf_build_report()` |
| เงื่อนไข 5 | ยอดติดลบแสดงในวงเล็บ | รูปแบบ `#,##0.00;(#,##0.00)` ทุกช่องจำนวนเงิน |
| เงื่อนไข 4 | จำกัดข้อมูลตามสิทธิ์สาขาของ User | `@br_code` ใน SP ทุกตัว · ดู **OI-03** |

## ส่วนที่ 3 — หัวรายงาน (แสดงทุกหน้า)

| CR | หัวข้อ | ที่มา |
| --- | --- | --- |
| 1 | ชื่อบริษัท | `t_comp` (Fixed Text, Bold) |
| 2 | สาขา : รหัส - ชื่อ | `c_branch` ← `h_br_code` + `h_br_name` (`f_branch_name()`) |
| 3 | ชื่อรายงาน | `t_title` (Fixed Text, Bold) |
| 4 | เลขที่ + วันที่ใบนำส่งฯ | `c_debitno` ← `h_debit_no_full`, `h_receipt_date` |
| 5 | ชื่อ + รหัสนายหน้า/ตัวแทน | `c_saleinfo` ← `h_sale_name`, `h_sale_ref` (`broker_code - sale_code`) |
| 6 | หน้าที่ : XX / YY | `c_page` ← `page() + ' / ' + pageCount()` |
| 7 | วันที่พิมพ์ | `c_prtdate` ← `f_nw_datetotdate( today() )` + `now()` |
| 8 | พิมพ์โดย | `c_prtby` ← `h_print_by` (`gs_user_name`) |

## ส่วนที่ 4 — Detail

| CR | หัวข้อ | คอลัมน์ / สูตร |
| --- | --- | --- |
| 9 | ค่าใช้จ่าย (%) — Grouping Policy/Endorse ตาม % | `group_caption` · `group(level=1 by=("group_caption"))` |
| 10 | จำนวนกรมธรรม์ | `policy_count` = `COUNT(DISTINCT policy_key)` |
| 11 | เบี้ยสุทธิ | `net_amount` |
| 12 | เบี้ยรวม | `total_amount` |
| — | เบี้ยประกันภัย ปัดเศษทศนิยม | `round_adj` — ดู **OI-01** |
| 13 | ค่าใช้จ่าย 1 | `ROUND(net × MIN(comrate, 0.18), 2)` |
| 14 | VAT 7% (ค่าใช้จ่าย 1) | `ROUND(exp1 × 0.07, 2)` เฉพาะตัวแทนจด VAT |
| 15 | TAX 3% (ค่าใช้จ่าย 1) | `ROUND(exp1 × 0.03, 2)` เฉพาะตัวแทนจด VAT |
| 16 | ค่าใช้จ่าย 2 | `ROUND(net × MAX(comrate − 0.18, 0), 2)` |
| 17 | VAT 7% (ค่าใช้จ่าย 2) | `ROUND(exp2 × 0.07, 2)` เฉพาะตัวแทนจด VAT |
| 18 | TAX 3% (ค่าใช้จ่าย 2) | `ROUND(exp2 × 0.03, 2)` เฉพาะตัวแทนจด VAT |
| 19 | ค่าใช้พิมพ์กรมธรรม์ | `ROUND(net × 0.01, 2)` เฉพาะ `flag_print_epolicy IN ('Y','1','Y/1')` |
| 20 | VAT 7% (ค่าพิมพ์ กธ.) | `ROUND(print_fee × 0.07, 2)` — ดู **OI-02** |
| 21 | TAX 3% (ค่าพิมพ์ กธ.) | `ROUND(print_fee × 0.03, 2)` — ดู **OI-02** |
| 22 | ภาษีหัก ณ ที่จ่าย (1%) — เฉพาะบรรทัดยอดรวม | `wht_1_percent` · แสดงเฉพาะแถว `ยอดรวม ภาคสมัครใจ` |
| 23 | ค่าดอกเบี้ย — เฉพาะบรรทัดยอดรวม | `interest_amount` · แสดงเฉพาะแถว `ยอดรวม ภาคสมัครใจ` |
| — | ยอดรวมต่อกลุ่ม "ยอดรวม : Policy 10.00%" | band `trailer.1` — `sum( x for group 1 )` |

## ส่วนที่ 5 — Footer (หน้าสุดท้าย)

| CR | หัวข้อ | ที่มา |
| --- | --- | --- |
| 24 | รวม Policy ภาคสมัครใจ | `sum( if( rec_type = 'P', x, 0 ) for all )` |
| 25 | รวม Endorse ภาคสมัครใจ | `sum( if( rec_type = 'E', x, 0 ) for all )` |
| 26 | ยอดรวม ภาคสมัครใจ | `sum( x for all )` |
| 27–32 | DD / OP / TS / BC / AC / TO OTHER | `f_dd`, `f_op`, `f_ts`, `f_bc`, `f_ac`, `f_to_other` ← `proc_acvun7123_header` |
| 33 | TOTAL AMOUNT DEBIT | `f_total_debit` |
| 34 | EXPENSE | `f_expense` |
| 35 | ยอดภาษีตัวแทน (Tax Deduct) | `f_tax_deduct` |
| 36 | ยอดชำระทั้งสิ้น | `f_total_net` (คำนวณใน `proc_acvun7123_header`) |
| 37 | ข้อความหัวหนังสือรับรองฯ | `t_wht_title` (Fixed Text) |
| 38 | ผู้มีหน้าที่หักภาษี ณ ที่จ่าย | `d_acvun7123_wht.payer_name` |
| 39 | ข้อความ "หนังสือรับรองการหักภาษี ณ ที่จ่าย" | หัวตารางของ `d_acvun7123_wht` |
| 40 | เล่มที่ | `d_acvun7123_wht.book_no` |
| 41 | เลขที่ | `d_acvun7123_wht.doc_no` |
| 42 | วันเดือนปี ภาษีที่จ่าย | `d_acvun7123_wht.c_tax_date` (พ.ศ.) |
| 43 | จำนวนเงิน ที่จ่าย | `d_acvun7123_wht.base_amount` |
| 44 | ภาษีที่หัก และนำส่งไว้ | `d_acvun7123_wht.tax_amount` |
| 45 | หมายเหตุ | `d_acvun7123_wht.remark` |
| 46 | รวมจำนวน N ฉบับ + ผลรวมข้อ 43, 44 | band `summary` ของ `d_acvun7123_wht` |
| 47 | ชื่อผู้ตัดชำระ (Operator Login) | `c_operator` ← `first( h_operator_login for all )` |

## สิ่งที่ยังไม่ครอบคลุมในการส่งมอบชุดนี้

| หัวข้อ | สถานะ |
| --- | --- |
| การผูกเมนู (CR ข้อ 1) | ต้องแก้ที่อ็อบเจกต์เมนูของระบบ ซึ่งไม่ได้อยู่ในชุดซอร์สที่ได้รับ — ดำเนินการตอน Import |
| ยืนยันชื่อตาราง/คอลัมน์ | ดู `docs/SCHEMA-CHECKLIST.md` — ต้องผ่าน DBA ก่อนรันสคริปต์ 03 และ 04 |
| การทดสอบกับข้อมูลจริง | ยังไม่ได้ทดสอบ — ไม่มีสิทธิ์เข้าถึงฐานข้อมูลในสภาพแวดล้อมนี้ |
