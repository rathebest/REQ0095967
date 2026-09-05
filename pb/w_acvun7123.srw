$PBExportHeader$w_acvun7123.srw
$PBExportComments$REQ0095968 - หน้าจอ สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย (ภาคสมัครใจ)
forward
global type w_acvun7123 from window
end type
type dw_cond from datawindow within w_acvun7123
end type
type dw_list from datawindow within w_acvun7123
end type
type st_status from statictext within w_acvun7123
end type
type cb_search from commandbutton within w_acvun7123
end type
type cb_selall from commandbutton within w_acvun7123
end type
type cb_selnone from commandbutton within w_acvun7123
end type
type cb_pdf from commandbutton within w_acvun7123
end type
type cb_excel from commandbutton within w_acvun7123
end type
type cb_exit from commandbutton within w_acvun7123
end type
end forward

global type w_acvun7123 from window
integer width = 4133
integer height = 3112
boolean titlebar = true
string title = "สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย - ภาคสมัครใจ"
boolean controlmenu = true
boolean minbox = true
boolean maxbox = true
boolean resizable = true
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
event ue_search ( )
event ue_print ( )
event ue_save ( )
event ue_exit ( )
dw_cond dw_cond
dw_list dw_list
st_status st_status
cb_search cb_search
cb_selall cb_selall
cb_selnone cb_selnone
cb_pdf cb_pdf
cb_excel cb_excel
cb_exit cb_exit
end type
global w_acvun7123 w_acvun7123

type variables
//  ---------------------------------------------------------------------------
//  REQ0095968 - สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย
//
//  ตัวแปร Global ของระบบที่หน้าจอนี้เรียกใช้ (ต้องมีอยู่แล้วในแอปพลิเคชัน)
//     gs_branch      - รหัสสาขาของ User ที่ Login (3 หลัก)
//     gs_user_name   - ชื่อ User ที่ Login (ใช้แสดง "พิมพ์โดย :")
//  ฟังก์ชัน Global ของระบบที่เรียกใช้
//     f_branch_name( string as_br_code )   - คืนชื่อสาขา
//     f_nw_datetotdate( date/datetime )    - แปลงเป็นวันที่แบบ พ.ศ. (ใช้ในนิพจน์ DataWindow)
//  ---------------------------------------------------------------------------
Transaction       itr_sql
DataStore         ids_sum       //  d_acvun7123_src_sum  - Detail ของรายงาน
DataStore         ids_hdr       //  d_acvun7123_src_hdr  - หัวรายงาน + กล่องสรุปท้ายรายงาน
DataStore         ids_rpt       //  d_acvun7123_rpt      - รูปเล่มรายงานสำหรับสั่งพิมพ์ PDF
DataStore         ids_xls       //  d_acvun7123_rpt_xls  - ตารางแบนสำหรับส่งออก Excel
DataWindowChild   idwc_oper     //  DropDownDataWindow ชื่อผู้ตัดชำระ

String            is_br_code    //  รหัสสาขาตามสิทธิ์ผู้ใช้ (แก้ไขไม่ได้)
String            is_br_name    //  ชื่อสาขา
String            is_out_path = "C:\Uxmotor\"
end variables

forward prototypes
public function integer wf_connect ()
public function integer wf_thai_to_datetime (string as_txt, ref datetime adt_out)
public function integer wf_search ()
public function long wf_selected_rows (ref long al_rows[])
public function integer wf_set_all (string as_flag)
public function integer wf_build_report (string as_debit_no, datastore ads_target)
public function integer wf_append_xls (string as_debit_no)
public function integer wf_print_pdf ()
public function integer wf_export_excel ()
end prototypes

public function integer wf_connect ();
//  สร้าง Transaction Object เฉพาะหน้าจอนี้ (ไม่ใช้ SQLCA หลัก) ตามมาตรฐานระบบ
itr_sql = CREATE Transaction

itr_sql.DBMS     = SQLCA.DBMS
itr_sql.Database = SQLCA.Database
itr_sql.ServerName = SQLCA.ServerName
itr_sql.LogId    = SQLCA.LogId
itr_sql.LogPass  = SQLCA.LogPass
itr_sql.AutoCommit = TRUE
itr_sql.DBParm   = SQLCA.DBParm

CONNECT USING itr_sql ;

IF itr_sql.SQLCode <> 0 THEN
	MessageBox( "Connect DB Error !", itr_sql.SQLErrText )
	DESTROY itr_sql
	RETURN -1
END IF

RETURN 1
end function

public function integer wf_thai_to_datetime (string as_txt, ref datetime adt_out);
//  แปลงข้อความวันที่แบบ วว/ดด/ปปปป (ปี พ.ศ.) ให้เป็น DateTime สำหรับส่งเข้า Stored Procedure
//  คืนค่า  1 = แปลงสำเร็จ, 0 = ผู้ใช้เว้นว่าง (adt_out = NULL), -1 = รูปแบบไม่ถูกต้อง
String   ls_txt
Integer  li_d, li_m, li_y
Date     ld_tmp

SetNull( adt_out )
ls_txt = Trim( as_txt )

//  EditMask อาจคืนค่าเป็นเครื่องหมายคั่นล้วนเมื่อผู้ใช้ยังไม่ได้กรอก
IF ls_txt = "" OR ls_txt = "/" OR ls_txt = "//" OR ls_txt = "00/00/0000" THEN
	RETURN 0
END IF

IF Len( ls_txt ) <> 10 THEN RETURN -1

li_d = Integer( Mid( ls_txt, 1, 2 ) )
li_m = Integer( Mid( ls_txt, 4, 2 ) )
li_y = Integer( Mid( ls_txt, 7, 4 ) )

IF li_d < 1 OR li_m < 1 OR li_y < 1 THEN RETURN -1

//  ปี พ.ศ. -> ค.ศ.  (รองรับกรณีผู้ใช้กรอกปี ค.ศ. มาโดยตรง)
IF li_y > 2400 THEN li_y = li_y - 543

ld_tmp = Date( li_y, li_m, li_d )
IF IsNull( ld_tmp ) OR String( ld_tmp, "yyyy-mm-dd" ) = "1900-01-01" THEN RETURN -1

adt_out = DateTime( ld_tmp, Time( "00:00:00" ) )
RETURN 1
end function

public function integer wf_search ();
//  อ่านเงื่อนไขจาก dw_cond ตรวจสอบความถูกต้อง แล้ว Retrieve รายการใบนำส่งฯ ลง dw_list
String    ls_oper, ls_dbn_fr, ls_dbn_to
DateTime  ldt_fr, ldt_to
Integer   li_fr, li_to
Long      ll_rows

dw_cond.AcceptText()

//  --- ชื่อผู้ตัดชำระ : ติ๊ก "เลือกรายการทั้งหมด" = ส่งค่าว่างให้ SP แปลว่าไม่กรอง ---
IF dw_cond.Object.flag_all_oper[1] = "1" THEN
	ls_oper = ""
ELSE
	ls_oper = Trim( dw_cond.Object.operator_login[1] )
	IF ls_oper = "" THEN
		MessageBox( "Warning !", "กรุณาระบุชื่อผู้ตัดชำระ หรือติ๊ก 'เลือกรายการทั้งหมด'" )
		RETURN -1
	END IF
END IF

//  --- ช่วงเลขที่ใบนำส่งฯ : ระบุมาข้างเดียวให้ถือเป็นใบเดียว ---
ls_dbn_fr = Trim( dw_cond.Object.debit_no_from[1] )
ls_dbn_to = Trim( dw_cond.Object.debit_no_to[1] )
IF ls_dbn_fr = "/" THEN ls_dbn_fr = ""
IF ls_dbn_to = "/" THEN ls_dbn_to = ""
IF ls_dbn_fr <> "" AND ls_dbn_to = "" THEN ls_dbn_to = ls_dbn_fr
IF ls_dbn_to <> "" AND ls_dbn_fr = "" THEN ls_dbn_fr = ls_dbn_to

//  --- ช่วงวันที่ใบนำส่งฯ ---
li_fr = wf_thai_to_datetime( dw_cond.Object.receipt_from[1], ldt_fr )
IF li_fr < 0 THEN
	MessageBox( "Warning !", "วันที่ใบนำส่งเงินชำระเบี้ย (จาก) ไม่ถูกต้อง กรุณาตรวจสอบ" )
	dw_cond.SetColumn( "receipt_from" )
	dw_cond.SetFocus()
	RETURN -1
END IF

li_to = wf_thai_to_datetime( dw_cond.Object.receipt_to[1], ldt_to )
IF li_to < 0 THEN
	MessageBox( "Warning !", "วันที่ใบนำส่งเงินชำระเบี้ย (ถึง) ไม่ถูกต้อง กรุณาตรวจสอบ" )
	dw_cond.SetColumn( "receipt_to" )
	dw_cond.SetFocus()
	RETURN -1
END IF

//  ระบุวันที่มาข้างเดียว ให้ถือเป็นวันเดียว
IF li_fr = 1 AND li_to = 0 THEN
	ldt_to = ldt_fr
	li_to  = 1
END IF
IF li_to = 1 AND li_fr = 0 THEN
	ldt_fr = ldt_to
	li_fr  = 1
END IF

IF li_fr = 1 AND li_to = 1 THEN
	IF Date( ldt_fr ) > Date( ldt_to ) THEN
		MessageBox( "Warning !", "ช่วงวันที่ใบนำส่งเงินชำระเบี้ยไม่ถูกต้อง กรุณาตรวจสอบ" )
		dw_cond.SetColumn( "receipt_from" )
		dw_cond.SetFocus()
		RETURN -1
	END IF
END IF

//  --- ต้องระบุอย่างน้อย 1 เงื่อนไข เพื่อไม่ให้ดึงข้อมูลทั้งสาขา ---
IF ls_dbn_fr = "" AND li_fr = 0 THEN
	MessageBox( "Warning !", "กรุณาระบุเลขที่ใบนำส่งเงินชำระเบี้ย หรือวันที่ใบนำส่งเงินชำระเบี้ย อย่างน้อย 1 เงื่อนไข" )
	RETURN -1
END IF

SetPointer( HourGlass! )
st_status.Text = "กำลังค้นหาข้อมูล ..."

ll_rows = dw_list.Retrieve( is_br_code, ls_oper, ls_dbn_fr, ls_dbn_to, ldt_fr, ldt_to )

IF ll_rows < 0 THEN
	st_status.Text = ""
	MessageBox( "Error !", "ค้นหาข้อมูลไม่สำเร็จ~r~n" + itr_sql.SQLErrText )
	RETURN -1
END IF

IF ll_rows = 0 THEN
	st_status.Text = "ไม่พบข้อมูล"
	MessageBox( "Not Found !", "ไม่พบข้อมูลใบนำส่งเงินชำระเบี้ยตามเงื่อนไขที่ระบุ" )
	RETURN 0
END IF

st_status.Text = "พบข้อมูล " + String( ll_rows, "#,##0" ) + " รายการ - กรุณาเลือกรายการที่ต้องการ แล้วสั่งพิมพ์ PDF หรือส่งออก Excel"
RETURN 1
end function

public function long wf_selected_rows (ref long al_rows[]);
//  รวบรวมหมายเลขแถวใน dw_list ที่ผู้ใช้ติ๊กเลือกไว้ (sel_flag = '1')
Long  ll_i, ll_cnt, ll_total
Long  ll_empty[]

al_rows = ll_empty
ll_cnt  = 0
ll_total = dw_list.RowCount()

dw_list.AcceptText()

FOR ll_i = 1 TO ll_total
	IF dw_list.Object.sel_flag[ll_i] = "1" THEN
		ll_cnt = ll_cnt + 1
		al_rows[ll_cnt] = ll_i
	END IF
NEXT

RETURN ll_cnt
end function

public function integer wf_set_all (string as_flag);
//  ติ๊ก / ยกเลิกการติ๊ก ทุกแถวใน dw_list
Long ll_i, ll_total

ll_total = dw_list.RowCount()
IF ll_total < 1 THEN RETURN 0

FOR ll_i = 1 TO ll_total
	dw_list.Object.sel_flag[ll_i] = as_flag
NEXT

RETURN 1
end function

public function integer wf_build_report (string as_debit_no, datastore ads_target);
//  สร้างรูปเล่มรายงานของใบนำส่งฯ 1 ใบ ลงใน ads_target (d_acvun7123_rpt)
//  คืนค่า จำนวนแถว Detail ที่ใส่ได้ (0 = ไม่มีข้อมูล, -1 = ผิดพลาด)
Long     ll_sum, ll_i, ll_new
String   ls_receipt, ls_sale_ref

ads_target.Reset()

IF ids_hdr.Retrieve( is_br_code, as_debit_no ) < 1 THEN RETURN -1
ll_sum = ids_sum.Retrieve( is_br_code, as_debit_no )
IF ll_sum < 0 THEN RETURN -1
IF ll_sum = 0 THEN RETURN 0

//  วันที่ใบนำส่งฯ แสดงเป็น พ.ศ.
IF IsNull( ids_hdr.Object.receipt_date[1] ) THEN
	ls_receipt = ""
ELSE
	ls_receipt = String( Date( ids_hdr.Object.receipt_date[1] ), "dd/mm/" ) &
	           + String( Year( Date( ids_hdr.Object.receipt_date[1] ) ) + 543 )
END IF

ls_sale_ref = Trim( ids_hdr.Object.broker_code[1] ) + " - " + Trim( ids_hdr.Object.sale_code[1] )

FOR ll_i = 1 TO ll_sum
	ll_new = ads_target.InsertRow( 0 )

	//  ----- ส่วน Detail (CR ข้อ 9 - 21) -----
	ads_target.Object.rec_type[ll_new]       = ids_sum.Object.rec_type[ll_i]
	ads_target.Object.comm_rate_pct[ll_new]  = ids_sum.Object.comm_rate_pct[ll_i]
	ads_target.Object.group_caption[ll_new]  = ids_sum.Object.group_caption[ll_i]
	ads_target.Object.product_group[ll_new]  = ids_sum.Object.product_group[ll_i]
	ads_target.Object.policy_count[ll_new]   = ids_sum.Object.policy_count[ll_i]
	ads_target.Object.net_amount[ll_new]     = ids_sum.Object.net_amount[ll_i]
	ads_target.Object.total_amount[ll_new]   = ids_sum.Object.total_amount[ll_i]
	ads_target.Object.round_adj[ll_new]      = ids_sum.Object.round_adj[ll_i]
	ads_target.Object.exp1[ll_new]           = ids_sum.Object.exp1[ll_i]
	ads_target.Object.exp1_vat7[ll_new]      = ids_sum.Object.exp1_vat7[ll_i]
	ads_target.Object.exp1_tax3[ll_new]      = ids_sum.Object.exp1_tax3[ll_i]
	ads_target.Object.exp2[ll_new]           = ids_sum.Object.exp2[ll_i]
	ads_target.Object.exp2_vat7[ll_new]      = ids_sum.Object.exp2_vat7[ll_i]
	ads_target.Object.exp2_tax3[ll_new]      = ids_sum.Object.exp2_tax3[ll_i]
	ads_target.Object.print_fee[ll_new]      = ids_sum.Object.print_fee[ll_i]
	ads_target.Object.print_vat7[ll_new]     = ids_sum.Object.print_vat7[ll_i]
	ads_target.Object.print_tax3[ll_new]     = ids_sum.Object.print_tax3[ll_i]

	//  ----- หัวรายงาน (CR ข้อ 2, 4, 5, 8, 47) -----
	ads_target.Object.h_br_code[ll_new]        = is_br_code
	ads_target.Object.h_br_name[ll_new]        = is_br_name
	ads_target.Object.h_debit_no[ll_new]       = as_debit_no
	ads_target.Object.h_debit_no_full[ll_new]  = ids_hdr.Object.debit_no_full[1]
	ads_target.Object.h_receipt_date[ll_new]   = ls_receipt
	ads_target.Object.h_sale_name[ll_new]      = ids_hdr.Object.sale_name[1]
	ads_target.Object.h_sale_ref[ll_new]       = ls_sale_ref
	ads_target.Object.h_operator_login[ll_new] = ids_hdr.Object.operator_login[1]
	ads_target.Object.h_print_by[ll_new]       = gs_user_name

	//  ----- CR ข้อ 22, 23 : แสดงบนบรรทัด "ยอดรวม ภาคสมัครใจ" -----
	ads_target.Object.wht_1_percent[ll_new]   = ids_hdr.Object.wht_1_percent[1]
	ads_target.Object.interest_amount[ll_new] = ids_hdr.Object.interest_amount[1]

	//  ----- CR ข้อ 27 - 36 : กล่องสรุปยอดท้ายรายงาน -----
	ads_target.Object.f_dd[ll_new]           = ids_hdr.Object.amt_dd[1]
	ads_target.Object.f_op[ll_new]           = ids_hdr.Object.amt_op[1]
	ads_target.Object.f_ts[ll_new]           = ids_hdr.Object.amt_ts[1]
	ads_target.Object.f_bc[ll_new]           = ids_hdr.Object.amt_bc[1]
	ads_target.Object.f_ac[ll_new]           = ids_hdr.Object.amt_ac[1]
	ads_target.Object.f_to_other[ll_new]     = ids_hdr.Object.amt_to_other[1]
	ads_target.Object.f_total_debit[ll_new]  = ids_hdr.Object.total_amount_debit[1]
	ads_target.Object.f_expense[ll_new]      = ids_hdr.Object.amt_expense[1]
	ads_target.Object.f_tax_deduct[ll_new]   = ids_hdr.Object.tax_deduct[1]
	ads_target.Object.f_total_net[ll_new]    = ids_hdr.Object.debitnote_total_net[1]
NEXT

//  เรียง Policy ก่อน Endorse, ภายในกลุ่มเรียง %ค่าใช้จ่ายจากน้อยไปมาก (CR ข้อ 3 / เงื่อนไขข้อ 6)
ads_target.SetSort( "rec_type A, comm_rate_pct A, product_group A" )
ads_target.Sort()
ads_target.GroupCalc()

RETURN ll_sum
end function

public function integer wf_append_xls (string as_debit_no);
//  ต่อท้ายข้อมูลของใบนำส่งฯ 1 ใบ ลงใน ids_xls (ตารางแบน - ใช้รวมทุกใบไว้ในไฟล์เดียว)
Long     ll_sum, ll_i, ll_new
String   ls_receipt, ls_sale_ref, ls_full, ls_oper, ls_sname
//  ตัวสะสมยอดรวม Policy / Endorse / ทั้งหมด (CR ข้อ 24 - 26)
Long        ld_cnt[3]
Decimal{2}  ld_net[3], ld_tot[3], ld_rnd[3]
Decimal{2}  ld_e1[3], ld_e1v[3], ld_e1t[3]
Decimal{2}  ld_e2[3], ld_e2v[3], ld_e2t[3]
Decimal{2}  ld_pf[3], ld_pfv[3], ld_pft[3]
Integer  li_k, li_b

IF ids_hdr.Retrieve( is_br_code, as_debit_no ) < 1 THEN RETURN -1
ll_sum = ids_sum.Retrieve( is_br_code, as_debit_no )
IF ll_sum < 1 THEN RETURN 0

IF IsNull( ids_hdr.Object.receipt_date[1] ) THEN
	ls_receipt = ""
ELSE
	ls_receipt = String( Date( ids_hdr.Object.receipt_date[1] ), "dd/mm/" ) &
	           + String( Year( Date( ids_hdr.Object.receipt_date[1] ) ) + 543 )
END IF

ls_full     = ids_hdr.Object.debit_no_full[1]
ls_sname    = ids_hdr.Object.sale_name[1]
ls_oper     = ids_hdr.Object.operator_login[1]
ls_sale_ref = Trim( ids_hdr.Object.broker_code[1] ) + " - " + Trim( ids_hdr.Object.sale_code[1] )

FOR li_k = 1 TO 3
	ld_cnt[li_k] = 0 ; ld_net[li_k] = 0 ; ld_tot[li_k] = 0 ; ld_rnd[li_k] = 0
	ld_e1[li_k]  = 0 ; ld_e1v[li_k] = 0 ; ld_e1t[li_k] = 0
	ld_e2[li_k]  = 0 ; ld_e2v[li_k] = 0 ; ld_e2t[li_k] = 0
	ld_pf[li_k]  = 0 ; ld_pfv[li_k] = 0 ; ld_pft[li_k] = 0
NEXT

FOR ll_i = 1 TO ll_sum
	ll_new = ids_xls.InsertRow( 0 )

	ids_xls.Object.debit_no_full[ll_new]     = ls_full
	ids_xls.Object.receipt_date_disp[ll_new] = ls_receipt
	ids_xls.Object.br_code[ll_new]           = is_br_code
	ids_xls.Object.br_name[ll_new]           = is_br_name
	ids_xls.Object.sale_ref[ll_new]          = ls_sale_ref
	ids_xls.Object.sale_name[ll_new]         = ls_sname
	ids_xls.Object.operator_login[ll_new]    = ls_oper
	ids_xls.Object.row_kind[ll_new]          = "D"

	ids_xls.Object.group_caption[ll_new]  = ids_sum.Object.group_caption[ll_i]
	ids_xls.Object.product_group[ll_new]  = ids_sum.Object.product_group[ll_i]
	ids_xls.Object.policy_count[ll_new]   = ids_sum.Object.policy_count[ll_i]
	ids_xls.Object.net_amount[ll_new]     = ids_sum.Object.net_amount[ll_i]
	ids_xls.Object.total_amount[ll_new]   = ids_sum.Object.total_amount[ll_i]
	ids_xls.Object.round_adj[ll_new]      = ids_sum.Object.round_adj[ll_i]
	ids_xls.Object.exp1[ll_new]           = ids_sum.Object.exp1[ll_i]
	ids_xls.Object.exp1_vat7[ll_new]      = ids_sum.Object.exp1_vat7[ll_i]
	ids_xls.Object.exp1_tax3[ll_new]      = ids_sum.Object.exp1_tax3[ll_i]
	ids_xls.Object.exp2[ll_new]           = ids_sum.Object.exp2[ll_i]
	ids_xls.Object.exp2_vat7[ll_new]      = ids_sum.Object.exp2_vat7[ll_i]
	ids_xls.Object.exp2_tax3[ll_new]      = ids_sum.Object.exp2_tax3[ll_i]
	ids_xls.Object.print_fee[ll_new]      = ids_sum.Object.print_fee[ll_i]
	ids_xls.Object.print_vat7[ll_new]     = ids_sum.Object.print_vat7[ll_i]
	ids_xls.Object.print_tax3[ll_new]     = ids_sum.Object.print_tax3[ll_i]

	//  สะสมยอด : 1 = Policy, 2 = Endorse, 3 = ทั้งหมด
	IF ids_sum.Object.rec_type[ll_i] = "P" THEN
		li_b = 1
	ELSE
		li_b = 2
	END IF

	FOR li_k = 1 TO 3
		IF li_k = 3 OR li_k = li_b THEN
			ld_cnt[li_k] = ld_cnt[li_k] + ids_sum.Object.policy_count[ll_i]
			ld_net[li_k] = ld_net[li_k] + ids_sum.Object.net_amount[ll_i]
			ld_tot[li_k] = ld_tot[li_k] + ids_sum.Object.total_amount[ll_i]
			ld_rnd[li_k] = ld_rnd[li_k] + ids_sum.Object.round_adj[ll_i]
			ld_e1[li_k]  = ld_e1[li_k]  + ids_sum.Object.exp1[ll_i]
			ld_e1v[li_k] = ld_e1v[li_k] + ids_sum.Object.exp1_vat7[ll_i]
			ld_e1t[li_k] = ld_e1t[li_k] + ids_sum.Object.exp1_tax3[ll_i]
			ld_e2[li_k]  = ld_e2[li_k]  + ids_sum.Object.exp2[ll_i]
			ld_e2v[li_k] = ld_e2v[li_k] + ids_sum.Object.exp2_vat7[ll_i]
			ld_e2t[li_k] = ld_e2t[li_k] + ids_sum.Object.exp2_tax3[ll_i]
			ld_pf[li_k]  = ld_pf[li_k]  + ids_sum.Object.print_fee[ll_i]
			ld_pfv[li_k] = ld_pfv[li_k] + ids_sum.Object.print_vat7[ll_i]
			ld_pft[li_k] = ld_pft[li_k] + ids_sum.Object.print_tax3[ll_i]
		END IF
	NEXT
NEXT

//  ----- บรรทัดยอดรวม 3 บรรทัด (CR ข้อ 24, 25, 26) -----
FOR li_k = 1 TO 3
	ll_new = ids_xls.InsertRow( 0 )

	ids_xls.Object.debit_no_full[ll_new]     = ls_full
	ids_xls.Object.receipt_date_disp[ll_new] = ls_receipt
	ids_xls.Object.br_code[ll_new]           = is_br_code
	ids_xls.Object.br_name[ll_new]           = is_br_name
	ids_xls.Object.sale_ref[ll_new]          = ls_sale_ref
	ids_xls.Object.sale_name[ll_new]         = ls_sname
	ids_xls.Object.operator_login[ll_new]    = ls_oper

	CHOOSE CASE li_k
		CASE 1
			ids_xls.Object.row_kind[ll_new]      = "P"
			ids_xls.Object.group_caption[ll_new] = "รวม Policy ภาคสมัครใจ"
		CASE 2
			ids_xls.Object.row_kind[ll_new]      = "E"
			ids_xls.Object.group_caption[ll_new] = "รวม Endorse ภาคสมัครใจ"
		CASE 3
			ids_xls.Object.row_kind[ll_new]      = "A"
			ids_xls.Object.group_caption[ll_new] = "ยอดรวม ภาคสมัครใจ"
	END CHOOSE

	ids_xls.Object.policy_count[ll_new] = ld_cnt[li_k]
	ids_xls.Object.net_amount[ll_new]   = ld_net[li_k]
	ids_xls.Object.total_amount[ll_new] = ld_tot[li_k]
	ids_xls.Object.round_adj[ll_new]    = ld_rnd[li_k]
	ids_xls.Object.exp1[ll_new]         = ld_e1[li_k]
	ids_xls.Object.exp1_vat7[ll_new]    = ld_e1v[li_k]
	ids_xls.Object.exp1_tax3[ll_new]    = ld_e1t[li_k]
	ids_xls.Object.exp2[ll_new]         = ld_e2[li_k]
	ids_xls.Object.exp2_vat7[ll_new]    = ld_e2v[li_k]
	ids_xls.Object.exp2_tax3[ll_new]    = ld_e2t[li_k]
	ids_xls.Object.print_fee[ll_new]    = ld_pf[li_k]
	ids_xls.Object.print_vat7[ll_new]   = ld_pfv[li_k]
	ids_xls.Object.print_tax3[ll_new]   = ld_pft[li_k]

	//  CR ข้อ 22, 23 และกล่องสรุปท้ายรายงาน แสดงเฉพาะบรรทัด "ยอดรวม ภาคสมัครใจ"
	IF li_k = 3 THEN
		ids_xls.Object.wht_1_percent[ll_new]   = ids_hdr.Object.wht_1_percent[1]
		ids_xls.Object.interest_amount[ll_new] = ids_hdr.Object.interest_amount[1]
		ids_xls.Object.f_dd[ll_new]            = ids_hdr.Object.amt_dd[1]
		ids_xls.Object.f_op[ll_new]            = ids_hdr.Object.amt_op[1]
		ids_xls.Object.f_ts[ll_new]            = ids_hdr.Object.amt_ts[1]
		ids_xls.Object.f_bc[ll_new]            = ids_hdr.Object.amt_bc[1]
		ids_xls.Object.f_ac[ll_new]            = ids_hdr.Object.amt_ac[1]
		ids_xls.Object.f_to_other[ll_new]      = ids_hdr.Object.amt_to_other[1]
		ids_xls.Object.f_total_debit[ll_new]   = ids_hdr.Object.total_amount_debit[1]
		ids_xls.Object.f_expense[ll_new]       = ids_hdr.Object.amt_expense[1]
		ids_xls.Object.f_tax_deduct[ll_new]    = ids_hdr.Object.tax_deduct[1]
		ids_xls.Object.f_total_net[ll_new]     = ids_hdr.Object.debitnote_total_net[1]
	END IF
NEXT

RETURN ll_sum
end function

public function integer wf_print_pdf ();
//  CR ข้อ 3 : พิมพ์ PDF A4 แนวนอน - สร้าง 1 ไฟล์ ต่อ 1 ใบนำส่งเงินชำระเบี้ย (แยกไฟล์)
Long     ll_rows[], ll_cnt, ll_i, ll_ok, ll_skip
String   ls_dbn, ls_file, ls_name

ll_cnt = wf_selected_rows( ll_rows )
IF ll_cnt < 1 THEN
	MessageBox( "Warning !", "กรุณาเลือกรายการใบนำส่งเงินชำระเบี้ยที่ต้องการพิมพ์ อย่างน้อย 1 รายการ" )
	RETURN -1
END IF

SetPointer( HourGlass! )
ll_ok   = 0
ll_skip = 0

FOR ll_i = 1 TO ll_cnt
	ls_dbn  = dw_list.Object.debit_no[ ll_rows[ll_i] ]
	ls_name = dw_list.Object.debit_no_full[ ll_rows[ll_i] ]

	st_status.Text = "กำลังสร้างเอกสาร " + String( ll_i ) + " / " + String( ll_cnt ) + " : " + ls_name

	IF wf_build_report( ls_dbn, ids_rpt ) < 1 THEN
		ll_skip = ll_skip + 1
		CONTINUE
	END IF

	//  1 ไฟล์ ต่อ 1 ใบนำส่งฯ - ตั้งชื่อจากคีย์จัดเก็บเพื่อเลี่ยงอักขระ '/' ในชื่อไฟล์
	ls_file = is_out_path + "W_ACVUN7123_" + is_br_code + "_" &
	        + Right( ls_dbn, 5 ) + "-" + Left( ls_dbn, 2 ) + ".pdf"

	ids_rpt.Modify( "DataWindow.Print.DocumentName='" + ls_name + "'" )
	ids_rpt.Modify( "DataWindow.Print.FileName='" + ls_file + "'" )

	IF ids_rpt.Print( FALSE ) = 1 THEN
		ll_ok = ll_ok + 1
	ELSE
		ll_skip = ll_skip + 1
	END IF
NEXT

st_status.Text = ""

IF ll_ok = 0 THEN
	MessageBox( "Error !", "สร้างเอกสารไม่สำเร็จ กรุณาตรวจสอบเครื่องพิมพ์และสิทธิ์การเขียนไฟล์ที่ " + is_out_path )
	RETURN -1
END IF

IF ll_skip > 0 THEN
	MessageBox( "Complete !", "สร้างเอกสารสำเร็จ " + String( ll_ok ) + " ไฟล์" + &
	            "~r~nไม่สำเร็จ / ไม่มีข้อมูล " + String( ll_skip ) + " รายการ~r~nที่ " + is_out_path )
ELSE
	MessageBox( "Complete !", "สร้างเอกสารเรียบร้อย " + String( ll_ok ) + " ไฟล์~r~nที่ " + is_out_path )
END IF

RETURN 1
end function

public function integer wf_export_excel ();
//  CR ข้อ 3 : ส่งออก Excel - รวมทุกใบนำส่งฯ ที่เลือกไว้ในไฟล์เดียว
Long     ll_rows[], ll_cnt, ll_i, ll_ok
String   ls_dbn, ls_file
Integer  li_rc

ll_cnt = wf_selected_rows( ll_rows )
IF ll_cnt < 1 THEN
	MessageBox( "Warning !", "กรุณาเลือกรายการใบนำส่งเงินชำระเบี้ยที่ต้องการส่งออก อย่างน้อย 1 รายการ" )
	RETURN -1
END IF

SetPointer( HourGlass! )
ids_xls.Reset()
ll_ok = 0

FOR ll_i = 1 TO ll_cnt
	ls_dbn = dw_list.Object.debit_no[ ll_rows[ll_i] ]
	st_status.Text = "กำลังรวบรวมข้อมูล " + String( ll_i ) + " / " + String( ll_cnt )
	IF wf_append_xls( ls_dbn ) > 0 THEN ll_ok = ll_ok + 1
NEXT

st_status.Text = ""

IF ids_xls.RowCount() < 1 THEN
	MessageBox( "Not Found !", "ไม่มีข้อมูลสำหรับส่งออกในรายการที่เลือก" )
	RETURN 0
END IF

ls_file = is_out_path + "W_ACVUN7123_" + is_br_code + ".xls"

//  หมายเหตุ : PowerBuilder 8.0 รองรับ Excel5! เป็นรูปแบบ Excel ที่สูงที่สุด
//             (Excel8! เริ่มมีใน PowerBuilder 10)
li_rc = ids_xls.SaveAs( ls_file, Excel5!, TRUE )

IF li_rc <> 1 THEN
	MessageBox( "Error !", "บันทึกข้อมูลไม่สำเร็จ~r~nกรุณาตรวจสอบสิทธิ์การเขียนไฟล์ที่ " + is_out_path )
	RETURN -1
END IF

MessageBox( "Complete !", "บันทึกไฟล์เรียบร้อย " + String( ll_ok ) + " ใบนำส่งฯ~r~nที่ " + ls_file )
RETURN 1
end function

on w_acvun7123.create
this.dw_cond = create dw_cond
this.dw_list = create dw_list
this.st_status = create st_status
this.cb_search = create cb_search
this.cb_selall = create cb_selall
this.cb_selnone = create cb_selnone
this.cb_pdf = create cb_pdf
this.cb_excel = create cb_excel
this.cb_exit = create cb_exit
this.Control[] = { this.dw_cond, &
this.dw_list, &
this.st_status, &
this.cb_search, &
this.cb_selall, &
this.cb_selnone, &
this.cb_pdf, &
this.cb_excel, &
this.cb_exit }
end on

on w_acvun7123.destroy
destroy(this.dw_cond)
destroy(this.dw_list)
destroy(this.st_status)
destroy(this.cb_search)
destroy(this.cb_selall)
destroy(this.cb_selnone)
destroy(this.cb_pdf)
destroy(this.cb_excel)
destroy(this.cb_exit)
end on

event open;
//  เตรียมการเชื่อมต่อ, ค่าเริ่มต้นของเงื่อนไข และ DataStore ที่ใช้ทั้งหน้าจอ
IF wf_connect() < 0 THEN
	Close( This )
	RETURN
END IF

//  สาขาตามสิทธิ์ผู้ใช้ที่ Login - แสดงอย่างเดียว แก้ไขไม่ได้
is_br_code = gs_branch
is_br_name = f_branch_name( is_br_code )

dw_cond.InsertRow( 0 )
dw_cond.Object.br_code[1]       = is_br_code
dw_cond.Object.br_name[1]       = is_br_name
dw_cond.Object.flag_all_oper[1] = "1"

//  DropDownDataWindow รายชื่อผู้ตัดชำระของสาขานี้
IF dw_cond.GetChild( "operator_login", idwc_oper ) = 1 THEN
	idwc_oper.SetTransObject( itr_sql )
	idwc_oper.Retrieve( is_br_code )
END IF

dw_list.SetTransObject( itr_sql )

ids_sum = CREATE DataStore
ids_sum.DataObject = "d_acvun7123_src_sum"
ids_sum.SetTransObject( itr_sql )

ids_hdr = CREATE DataStore
ids_hdr.DataObject = "d_acvun7123_src_hdr"
ids_hdr.SetTransObject( itr_sql )

//  ids_rpt ต้องมี Transaction Object ด้วย เพราะมี Nested Report (d_acvun7123_wht)
//  ที่ต้อง Retrieve เองตอนสั่งพิมพ์
ids_rpt = CREATE DataStore
ids_rpt.DataObject = "d_acvun7123_rpt"
ids_rpt.SetTransObject( itr_sql )

ids_xls = CREATE DataStore
ids_xls.DataObject = "d_acvun7123_rpt_xls"

st_status.Text = "กรุณาระบุเงื่อนไข แล้วกดปุ่ม ค้นหา"
end event

event close;
//  คืนทรัพยากรทั้งหมด : ป้องกัน Memory Leak และ Orphan Connection
IF IsValid( ids_sum ) THEN DESTROY ids_sum
IF IsValid( ids_hdr ) THEN DESTROY ids_hdr
IF IsValid( ids_rpt ) THEN DESTROY ids_rpt
IF IsValid( ids_xls ) THEN DESTROY ids_xls

IF IsValid( itr_sql ) THEN
	DISCONNECT USING itr_sql ;
	DESTROY itr_sql
END IF
end event

event ue_search;
wf_search()
end event

event ue_print;
wf_print_pdf()
end event

event ue_save;
wf_export_excel()
end event

event ue_exit;
Close( This )
end event

type dw_cond from datawindow within w_acvun7123
integer x = 18
integer y = 16
integer width = 4060
integer height = 360
integer taborder = 10
string dataobject = "d_acvun7123_cond"
boolean border = true
borderstyle borderstyle = stylelowered!
boolean livescroll = true
end type

event itemchanged;
//  ติ๊ก "เลือกรายการทั้งหมด" -> ล้างและล็อกช่องชื่อผู้ตัดชำระ
IF dwo.Name = "flag_all_oper" THEN
	IF data = "1" THEN
		This.Object.operator_login[row] = ""
	END IF
END IF
end event


type dw_list from datawindow within w_acvun7123
integer x = 18
integer y = 392
integer width = 4060
integer height = 2140
integer taborder = 20
string dataobject = "d_acvun7123_list"
boolean hscrollbar = true
boolean vscrollbar = true
boolean livescroll = true
boolean border = true
borderstyle borderstyle = stylelowered!
end type

event clicked;
//  คลิกที่แถวใดก็ได้เพื่อสลับสถานะการเลือก (นอกเหนือจากการติ๊กที่ช่อง Checkbox)
IF row < 1 THEN RETURN

IF dwo.Name <> "sel_flag" THEN
	IF This.Object.sel_flag[row] = "1" THEN
		This.Object.sel_flag[row] = "0"
	ELSE
		This.Object.sel_flag[row] = "1"
	END IF
END IF
end event

type st_status from statictext within w_acvun7123
integer x = 18
integer y = 2552
integer width = 4060
integer height = 76
boolean enabled = false
string text = ""
long backcolor = 67108864
boolean focusrectangle = false
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

type cb_search from commandbutton within w_acvun7123
integer x = 18
integer y = 2640
integer width = 402
integer height = 110
integer taborder = 30
string text = "ค้นหา"
boolean default = true
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
Parent.TriggerEvent( "ue_search" )
end event

type cb_selall from commandbutton within w_acvun7123
integer x = 434
integer y = 2640
integer width = 402
integer height = 110
integer taborder = 40
string text = "เลือกทั้งหมด"
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
wf_set_all( "1" )
end event

type cb_selnone from commandbutton within w_acvun7123
integer x = 850
integer y = 2640
integer width = 402
integer height = 110
integer taborder = 50
string text = "ยกเลิกทั้งหมด"
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
wf_set_all( "0" )
end event

type cb_pdf from commandbutton within w_acvun7123
integer x = 2606
integer y = 2640
integer width = 466
integer height = 110
integer taborder = 60
string text = "พิมพ์ PDF"
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
Parent.TriggerEvent( "ue_print" )
end event

type cb_excel from commandbutton within w_acvun7123
integer x = 3086
integer y = 2640
integer width = 466
integer height = 110
integer taborder = 70
string text = "ส่งออก Excel"
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
Parent.TriggerEvent( "ue_save" )
end event

type cb_exit from commandbutton within w_acvun7123
integer x = 3566
integer y = 2640
integer width = 402
integer height = 110
integer taborder = 80
string text = "ออก"
boolean cancel = true
integer textsize = -9
string facename = "Tahoma"
integer weight = 400
fontcharset fontcharset = thaiansi!
end type

event clicked;
Parent.TriggerEvent( "ue_exit" )
end event
