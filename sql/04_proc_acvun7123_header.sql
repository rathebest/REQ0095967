/* =============================================================================
   CR         : REQ0095968 - สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย
   Object     : [dbo].[proc_acvun7123_header] (Stored Procedure - NEW)
                [dbo].[proc_acvun7123_wht]    (Stored Procedure - NEW)
   Purpose    : proc_acvun7123_header -> ข้อมูลหัวรายงาน (CR ข้อ 2, 4, 5, 47)
                                         และกล่องสรุปท้ายรายงาน (CR ข้อ 22, 23, 27 - 36)
                proc_acvun7123_wht    -> รายการหนังสือรับรองการหักภาษี ณ ที่จ่าย
                                         (CR ข้อ 37 - 46)
   Used by    : w_acvun7123.wf_build_report()
   Target DB  : ฐานข้อมูลระบบงานภาคสมัครใจ (VMI)
   ---------------------------------------------------------------------------
   !! ต้องยืนยันกับ DBA ก่อนติดตั้ง (ดู docs/SCHEMA-CHECKLIST.md) !!
      คอลัมน์ในบล็อก [SCHEMA BINDING] ยังไม่ได้ตรวจสอบกับ Data Dictionary จริง
   ============================================================================= */
IF OBJECT_ID('dbo.proc_acvun7123_header') IS NOT NULL
    DROP PROCEDURE dbo.proc_acvun7123_header
GO

CREATE PROCEDURE [dbo].[proc_acvun7123_header]
(
    @br_code  CHAR(3),
    @debit_no CHAR(7)
)
AS
BEGIN
    SET NOCOUNT ON

    SELECT
        /* --- CR ข้อ 2 : สาขา (ชื่อสาขาให้ PowerScript เติมด้วย f_branch_name()) --- */
        dl.br_code,

        /* --- CR ข้อ 4 : เลขที่ / วันที่ ใบนำส่งเงินชำระเบี้ย --- */
        dl.debit_no,
        debit_no_disp = RIGHT(RTRIM(dl.debit_no), 5) + '/' + LEFT(dl.debit_no, 2),
        debit_no_full = 'V' + dl.debit_no_belongto + dl.flag_debit_type
                            + RIGHT(RTRIM(dl.debit_no), 5) + '/' + LEFT(dl.debit_no, 2),
        dl.receipt_date,

        /* --- CR ข้อ 5 : ชื่อ / รหัส นายหน้า - ตัวแทน --- */
        dl.sale_code,
        s.broker_code,
        sale_name = LTRIM(RTRIM(ISNULL(p.prename_desc, '')))
                  + LTRIM(RTRIM(ISNULL(s.fname, '')))
                  + ' ' + LTRIM(RTRIM(ISNULL(s.lname, ''))),
        sale_vat_registered = CASE WHEN ISNULL(s.flag_vat, 'N') IN ('Y', '1')
                                   THEN 'Y' ELSE 'N' END,

        /* --- CR ข้อ 47 : ชื่อผู้ตัดชำระ (Operator Login) --- */
        dl.operator_login,

        /* ---------------------------------------------------------------
           [SCHEMA BINDING] CR ข้อ 22, 23 - แสดงเฉพาะบรรทัด "ยอดรวม ภาคสมัครใจ"
           --------------------------------------------------------------- */
        wht_1_percent    = ISNULL(dl.vmi_wt_amount, 0),        -- ภาษีหัก ณ ที่จ่าย (1%)
        interest_amount  = ISNULL(dl.total_term_amount, 0),    -- ค่าดอกเบี้ย

        /* ---------------------------------------------------------------
           [SCHEMA BINDING] CR ข้อ 27 - 35 : กล่องสรุปยอดท้ายรายงาน
           --------------------------------------------------------------- */
        amt_dd           = ISNULL(dl.dd_amount, 0),                       -- 27 DD
        amt_op           = ISNULL(dl.total_amount_deposit, 0),            -- 28 OP
        amt_ts           = ISNULL(dl.amt_return_to_salemgr, 0),           -- 29 TS
        amt_bc           = ISNULL(dl.bc_amount, 0),                       -- 30 BC
        amt_ac           = ISNULL(dl.total_amount_ac, 0),                 -- 31 AC
        amt_to_other     = ISNULL(dl.amt_return_to_cmi, 0),               -- 32 TO OTHER
        total_amount_debit = ISNULL(dl.total_amount_debit, 0),            -- 33 TOTAL AMOUNT DEBIT
        amt_expense      = ISNULL(dl.total_amount_br_income_expense, 0),  -- 34 EXPENSE
        tax_deduct       = ISNULL(dl.tax_deduct, 0),                      -- 35 ยอดภาษีตัวแทน

        /* ---------------------------------------------------------------
           CR ข้อ 36 : ยอดชำระทั้งสิ้น (DEBITNOTE TOTAL NET)
           = TOTAL AMOUNT DEBIT หักด้วยรายการนำมาหักทุกก้อน
           ตรวจสอบกับตัวอย่างในไฟล์เงื่อนไข : 95,581.89 - 18,241.46 = 77,340.43
           --------------------------------------------------------------- */
        debitnote_total_net =
              ISNULL(dl.total_amount_debit, 0)
            - ISNULL(dl.dd_amount, 0)
            - ISNULL(dl.total_amount_deposit, 0)
            - ISNULL(dl.amt_return_to_salemgr, 0)
            - ISNULL(dl.bc_amount, 0)
            - ISNULL(dl.total_amount_ac, 0)
            - ISNULL(dl.amt_return_to_cmi, 0)
            - ISNULL(dl.tax_deduct, 0)
    FROM dbo.debit_list dl
    LEFT JOIN dbo.sale    s ON s.sale_code = dl.sale_code
    LEFT JOIN dbo.prename p ON p.code      = s.code
    WHERE dl.br_code         = @br_code
      AND dl.debit_no        = @debit_no
      AND dl.debit_undtype   = 'V'
      AND dl.flag_debit_type = 'C'
END
GO

GRANT EXECUTE ON dbo.proc_acvun7123_header TO PUBLIC
GO


/* =============================================================================
   [dbo].[proc_acvun7123_wht]
   CR ข้อ 37 - 46 : หนังสือรับรองการหักภาษี ณ ที่จ่าย จากเงินได้พึงประเมินประเภท
   ค่าเบี้ยประกันวินาศภัย - รถยนต์ภาคสมัครใจ
   ข้อมูลมาจากหน้าจอ "บันทึกใบนำส่งเงิน - รับชำระเบี้ย" ตามเลขที่ใบนำส่งเงิน
   ============================================================================= */
IF OBJECT_ID('dbo.proc_acvun7123_wht') IS NOT NULL
    DROP PROCEDURE dbo.proc_acvun7123_wht
GO

CREATE PROCEDURE [dbo].[proc_acvun7123_wht]
(
    @br_code  CHAR(3),
    @debit_no CHAR(7)
)
AS
BEGIN
    SET NOCOUNT ON

    /* [SCHEMA BINDING] ตาราง dbo.debit_wht = รายการหนังสือรับรองฯ ต่อ 1 ใบนำส่งฯ */
    SELECT
        seq_no       = w.seq_no,
        payer_name   = LTRIM(RTRIM(ISNULL(w.payer_name, ''))),  -- 38 ผู้มีหน้าที่หักภาษี ณ ที่จ่าย
        book_no      = LTRIM(RTRIM(ISNULL(w.book_no, ''))),     -- 40 เล่มที่
        doc_no       = LTRIM(RTRIM(ISNULL(w.doc_no, ''))),      -- 41 เลขที่
        tax_date     = w.tax_date,                              -- 42 วันเดือนปี ภาษีที่จ่าย
        base_amount  = ISNULL(w.base_amount, 0),                -- 43 จำนวนเงินที่จ่าย
        tax_amount   = ISNULL(w.tax_amount, 0),                 -- 44 ภาษีที่หักและนำส่งไว้
        remark       = LTRIM(RTRIM(ISNULL(w.remark, '')))       -- 45 หมายเหตุ
    FROM dbo.debit_wht w
    WHERE w.br_code       = @br_code
      AND w.debit_no      = @debit_no
      AND w.debit_undtype = 'V'
    ORDER BY w.seq_no ASC
END
GO

GRANT EXECUTE ON dbo.proc_acvun7123_wht TO PUBLIC
GO
