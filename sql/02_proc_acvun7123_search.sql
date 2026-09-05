/* =============================================================================
   CR         : REQ0095968 - สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย
   Object     : [dbo].[proc_acvun7123_search]        (Stored Procedure - NEW)
                [dbo].[proc_acvun7123_operator_list] (Stored Procedure - NEW)
   Purpose    : ค้นหารายการใบนำส่งเงินชำระเบี้ย (debit note) ภาคสมัครใจ
                ตามเงื่อนไขบนหน้าจอหลัก w_acvun7123 เพื่อนำไปแสดงใน DataWindow
                d_acvun7123_list ให้ผู้ใช้เลือก (ติ๊ก) ได้ตั้งแต่ 1 รายการขึ้นไป
                ก่อนสั่งพิมพ์ PDF หรือส่งออก Excel
   Used by    : w_acvun7123.wf_search()  ->  d_acvun7123_list
   Target DB  : ฐานข้อมูลระบบงานภาคสมัครใจ (VMI)
   Depends on : [dbo].[fn_acvun_debitno_key] (สคริปต์ 01)
   ---------------------------------------------------------------------------
   หมายเหตุการกรองข้อมูล (อ้างอิงหน้าจอเดิม W_ACVUN711 / d_acvun7122)
     debit_undtype   = 'V'   -> เฉพาะงานภาคสมัครใจ (Voluntary)
     flag_debit_type = 'C'   -> เฉพาะใบนำส่งประเภทตัดชำระเบี้ย
     br_code         = @br_code (สาขาตามสิทธิ์ผู้ใช้ที่ Login - แก้ไขไม่ได้บนหน้าจอ)
   ============================================================================= */
IF OBJECT_ID('dbo.proc_acvun7123_search') IS NOT NULL
    DROP PROCEDURE dbo.proc_acvun7123_search
GO

CREATE PROCEDURE [dbo].[proc_acvun7123_search]
(
    @br_code        CHAR(3),                  -- รหัสสาขา (บังคับ - มาจาก Login)
    @operator_login VARCHAR(15) = NULL,       -- ชื่อผู้ตัดชำระ; NULL/'' = ทุกคน (เลือกทั้งหมด)
    @debit_no_from  VARCHAR(20) = NULL,       -- เลขใบนำส่งฯ From 'NNNNN/YY'; NULL/'' = ไม่จำกัด
    @debit_no_to    VARCHAR(20) = NULL,       -- เลขใบนำส่งฯ To   'NNNNN/YY'; NULL/'' = ไม่จำกัด
    @receipt_from   DATETIME    = NULL,       -- วันที่ใบนำส่งฯ From; NULL = ไม่จำกัด
    @receipt_to     DATETIME    = NULL        -- วันที่ใบนำส่งฯ To;   NULL = ไม่จำกัด
)
AS
BEGIN
    SET NOCOUNT ON

    DECLARE @key_from CHAR(7),
            @key_to   CHAR(7),
            @dt_from  DATETIME,
            @dt_to    DATETIME

    /* ---------------------------------------------------------------------
       1) แปลงช่วงเลขที่ใบนำส่งฯ ให้เป็นคีย์จัดเก็บ YYNNNNN
          ถ้าผู้ใช้ไม่ระบุ (หรือระบุรูปแบบไม่ถูกต้อง) ให้เปิดช่วงกว้างสุด
       --------------------------------------------------------------------- */
    SET @key_from = dbo.fn_acvun_debitno_key(@debit_no_from)
    SET @key_to   = dbo.fn_acvun_debitno_key(@debit_no_to)

    IF @key_from IS NULL SET @key_from = '0000000'
    IF @key_to   IS NULL SET @key_to   = '9999999'

    /* ---------------------------------------------------------------------
       2) ช่วงวันที่ใบนำส่งฯ
          ตัดเวลาออกจากขอบล่าง และขยายขอบบนถึงสิ้นวัน เพื่อให้ receipt_date
          ที่มีเวลาติดมาด้วยถูกรวมอยู่ในช่วงที่ผู้ใช้ระบุครบถ้วน
       --------------------------------------------------------------------- */
    SET @dt_from = CONVERT(DATETIME,
                       CONVERT(CHAR(8), ISNULL(@receipt_from,
                           CONVERT(DATETIME, '17530101', 112)), 112), 112)

    IF @receipt_to IS NULL
        /* ขอบบนสุดของชนิด DATETIME - ห้ามใช้ DATEADD(dd,1,'9999-12-31') เพราะจะ overflow */
        SET @dt_to = CONVERT(DATETIME, '9999-12-31 23:59:59.997', 121)
    ELSE
        SET @dt_to = DATEADD(ms, -3, DATEADD(dd, 1,
                         CONVERT(DATETIME, CONVERT(CHAR(8), @receipt_to, 112), 112)))

    /* ---------------------------------------------------------------------
       3) รายการใบนำส่งเงินชำระเบี้ยที่ตรงเงื่อนไข
          เรียงตามเลขที่ใบนำส่งฯ จากน้อยไปมาก (CR ข้อ 3 - Sorting)
       --------------------------------------------------------------------- */
    SELECT
        /* --- ช่องติ๊กเลือกบน d_acvun7123_list (ผู้ใช้แก้ไขได้, ค่าเริ่มต้น = ไม่เลือก) --- */
        sel_flag = CONVERT(CHAR(1), '0'),

        /* --- คีย์และการแสดงผลเลขที่ใบนำส่งฯ --- */
        dl.debit_no,                                         -- คีย์จัดเก็บ YYNNNNN
        debit_no_disp  = RIGHT(RTRIM(dl.debit_no), 5) + '/' + LEFT(dl.debit_no, 2),
        debit_no_full  = 'V' + dl.debit_no_belongto + dl.flag_debit_type
                             + RIGHT(RTRIM(dl.debit_no), 5) + '/' + LEFT(dl.debit_no, 2),
        dl.debit_no_belongto,
        dl.flag_debit_type,

        /* --- สาขา / วันที่ / ผู้ตัดชำระ --- */
        dl.br_code,
        dl.receipt_date,
        dl.operator_login,

        /* --- นายหน้า / ตัวแทน --- */
        dl.sale_code,
        s.broker_code,
        sale_name = LTRIM(RTRIM(ISNULL(p.prename_desc, '')))
                  + LTRIM(RTRIM(ISNULL(s.fname, '')))
                  + ' ' + LTRIM(RTRIM(ISNULL(s.lname, ''))),

        /* --- ยอดเงินระดับใบนำส่งฯ (ใช้แสดงบน List ให้ผู้ใช้ตัดสินใจเลือก) --- */
        dl.total_amount_debit,
        dl.total_amount_deposit,
        dl.total_amount_ac,
        dl.total_amount_br_income_expense,
        dl.total_amount_return,
        dl.amt_return_to_salemgr,
        dl.amt_return_to_cmi
    FROM dbo.debit_list dl
    LEFT JOIN dbo.sale    s ON s.sale_code = dl.sale_code
    LEFT JOIN dbo.prename p ON p.code      = s.code
    WHERE dl.br_code         = @br_code
      AND dl.debit_undtype   = 'V'
      AND dl.flag_debit_type = 'C'
      AND dl.debit_no      BETWEEN @key_from AND @key_to
      AND dl.receipt_date  BETWEEN @dt_from  AND @dt_to
      AND (@operator_login IS NULL
           OR LTRIM(RTRIM(@operator_login)) = ''
           OR dl.operator_login = @operator_login)
    ORDER BY dl.debit_no ASC
END
GO

GRANT EXECUTE ON dbo.proc_acvun7123_search TO PUBLIC
GO


/* =============================================================================
   [dbo].[proc_acvun7123_operator_list]
   คืนรายชื่อผู้ตัดชำระ (operator_login) ที่มีใบนำส่งฯ ภาคสมัครใจอยู่จริงในสาขา
   ใช้เป็นแหล่งข้อมูลของ DropDownDataWindow ช่อง "ชื่อผู้ตัดชำระ" บนหน้าจอค้นหา
   ============================================================================= */
IF OBJECT_ID('dbo.proc_acvun7123_operator_list') IS NOT NULL
    DROP PROCEDURE dbo.proc_acvun7123_operator_list
GO

CREATE PROCEDURE [dbo].[proc_acvun7123_operator_list]
(
    @br_code CHAR(3)
)
AS
BEGIN
    SET NOCOUNT ON

    SELECT DISTINCT
        operator_login = dl.operator_login
    FROM dbo.debit_list dl
    WHERE dl.br_code         = @br_code
      AND dl.debit_undtype   = 'V'
      AND dl.flag_debit_type = 'C'
      AND dl.operator_login IS NOT NULL
      AND LTRIM(RTRIM(dl.operator_login)) <> ''
    ORDER BY dl.operator_login ASC
END
GO

GRANT EXECUTE ON dbo.proc_acvun7123_operator_list TO PUBLIC
GO
