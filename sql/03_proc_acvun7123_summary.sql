/* =============================================================================
   CR         : REQ0095968 - สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย
   Object     : [dbo].[proc_acvun7123_summary]   (Stored Procedure - NEW)
   Purpose    : คืนข้อมูลส่วน Detail ของรายงาน (CR ข้อ 9 - 23) สำหรับใบนำส่งฯ 1 ใบ
                จัดกลุ่มเป็น Policy / Endorse -> %ค่าใช้จ่าย -> กลุ่มผลิตภัณฑ์
                โดยเรียง %ค่าใช้จ่าย จากน้อยไปมาก (CR ข้อ 3 และเงื่อนไขข้อ 6)
   Used by    : w_acvun7123.wf_build_report()  ->  d_acvun7123_rpt / _rpt_xls
   Target DB  : ฐานข้อมูลระบบงานภาคสมัครใจ (VMI)
   ---------------------------------------------------------------------------
   สูตรคำนวณ (ตามเอกสาร CR ข้อ 13 - 21 และไฟล์เงื่อนไขรายงาน)
     ค่าใช้จ่าย 1         = เบี้ยสุทธิ x %ค่าใช้จ่าย เฉพาะส่วนที่ไม่เกินอัตรา คปภ. 18%
     ค่าใช้จ่าย 2         = เบี้ยสุทธิ x %ค่าใช้จ่าย เฉพาะส่วนที่เกิน 18%
     ค่าใช้พิมพ์กรมธรรม์  = เบี้ยสุทธิ x 1%  เฉพาะกรมธรรม์ที่ตัวแทนพิมพ์เอง (Y/1)
     VAT 7% / TAX 3%     = คิดจากค่าใช้จ่ายแต่ละก้อน เฉพาะนายหน้า/ตัวแทนที่จดทะเบียน VAT
     ทุกยอดปัดทศนิยม 2 ตำแหน่งด้วย ROUND(x, 2)
   ---------------------------------------------------------------------------
   !! ต้องยืนยันกับ DBA ก่อนติดตั้ง (ดู docs/SCHEMA-CHECKLIST.md) !!
      ชื่อตาราง/คอลัมน์ในบล็อก [SCHEMA BINDING] ด้านล่าง อ้างอิงจากเอกสาร CR และ
      หน้าจอเดิม d_acvun7122 เท่านั้น ยังไม่ได้ตรวจสอบกับ Data Dictionary จริง
   ============================================================================= */
IF OBJECT_ID('dbo.proc_acvun7123_summary') IS NOT NULL
    DROP PROCEDURE dbo.proc_acvun7123_summary
GO

CREATE PROCEDURE [dbo].[proc_acvun7123_summary]
(
    @br_code   CHAR(3),          -- รหัสสาขาเจ้าของใบนำส่งฯ
    @debit_no  CHAR(7)           -- เลขที่ใบนำส่งฯ รูปแบบจัดเก็บ YYNNNNN
)
AS
BEGIN
    SET NOCOUNT ON

    DECLARE @vat_registered BIT,
            @sale_code      CHAR(5)

    /* ---------------------------------------------------------------------
       0) นายหน้า/ตัวแทนของใบนำส่งฯ ใบนี้ + สถานะจดทะเบียน VAT
          VAT 7% และ TAX 3% จะคิดเฉพาะตัวแทนที่จดทะเบียน VAT เท่านั้น
       --------------------------------------------------------------------- */
    SELECT @sale_code = dl.sale_code
    FROM dbo.debit_list dl
    WHERE dl.br_code         = @br_code
      AND dl.debit_no        = @debit_no
      AND dl.debit_undtype   = 'V'
      AND dl.flag_debit_type = 'C'

    SELECT @vat_registered =
           CASE WHEN ISNULL(s.flag_vat, 'N') IN ('Y', '1') THEN 1 ELSE 0 END
    FROM dbo.sale s
    WHERE s.sale_code = @sale_code

    SET @vat_registered = ISNULL(@vat_registered, 0)

    /* =====================================================================
       [SCHEMA BINDING]
       1) ดึงรายการเบี้ยระดับบรรทัด (กรมธรรม์ / สลักหลัง) ที่ผูกกับใบนำส่งฯ ใบนี้
          - rec_type       'P' = Policy, 'E' = Endorse
          - sign_factor    สลักหลังลดเบี้ยให้เป็นค่าติดลบ (รายงานแสดงในวงเล็บ)
          - comm_rate      %ค่าใช้จ่าย เก็บเป็นทศนิยม เช่น 0.18 = 18%
       ===================================================================== */
    CREATE TABLE #line
    (
        rec_type        CHAR(1)        NOT NULL,
        comm_rate       DECIMAL(9, 6)  NOT NULL,
        product_group   VARCHAR(10)    NOT NULL,
        policy_key      VARCHAR(30)    NOT NULL,
        net             DECIMAL(18, 2) NOT NULL,
        total           DECIMAL(18, 2) NOT NULL,
        round_adj       DECIMAL(18, 2) NOT NULL,
        print_epolicy   CHAR(3)        NOT NULL
    )

    /* ---- 1.1 กรมธรรม์ (Policy) ---- */
    INSERT INTO #line
        (rec_type, comm_rate, product_group, policy_key, net, total, round_adj, print_epolicy)
    SELECT
        'P',
        ISNULL(pol.comrate, 0),
        ISNULL(pol.product_group, ''),
        RTRIM(pol.pol_yr) + RTRIM(pol.pol_br) + RTRIM(pol.pol_no) + RTRIM(pol.pol_pre),
        ISNULL(pol.net, 0),
        ISNULL(pol.total, 0),
        ISNULL(pol.round_adj, 0),
        ISNULL(pol.flag_print_epolicy, 'N')
    FROM dbo.debit_detail dd
    INNER JOIN dbo.policy pol
            ON  pol.pol_yr  = dd.pol_yr
            AND pol.pol_br  = dd.pol_br
            AND pol.pol_no  = dd.pol_no
            AND pol.pol_pre = dd.pol_pre
    WHERE dd.br_code         = @br_code
      AND dd.debit_no        = @debit_no
      AND dd.debit_undtype   = 'V'
      AND dd.flag_debit_type = 'C'
      AND (dd.endos_no IS NULL OR dd.endos_no IN ('', '000000'))

    /* ---- 1.2 สลักหลัง (Endorse) : เพิ่มเบี้ย = บวก, ลดเบี้ย = ลบ ---- */
    INSERT INTO #line
        (rec_type, comm_rate, product_group, policy_key, net, total, round_adj, print_epolicy)
    SELECT
        'E',
        ISNULL(pol.comrate, 0),
        ISNULL(pol.product_group, ''),
        RTRIM(en.endos_yr) + RTRIM(en.endos_no),
        ISNULL(en.net,       0) * CASE WHEN en.flag_inc_dec = 'D' THEN -1 ELSE 1 END,
        ISNULL(en.total,     0) * CASE WHEN en.flag_inc_dec = 'D' THEN -1 ELSE 1 END,
        ISNULL(en.round_adj, 0) * CASE WHEN en.flag_inc_dec = 'D' THEN -1 ELSE 1 END,
        ISNULL(pol.flag_print_epolicy, 'N')
    FROM dbo.debit_detail dd
    INNER JOIN dbo.endorse en
            ON  en.endos_yr = dd.endos_yr
            AND en.endos_no = dd.endos_no
    INNER JOIN dbo.policy pol
            ON  pol.pol_yr  = en.pol_yr
            AND pol.pol_br  = en.pol_br
            AND pol.pol_no  = en.pol_no
            AND pol.pol_pre = en.pol_pre
    WHERE dd.br_code         = @br_code
      AND dd.debit_no        = @debit_no
      AND dd.debit_undtype   = 'V'
      AND dd.flag_debit_type = 'C'
      AND dd.endos_no IS NOT NULL
      AND dd.endos_no NOT IN ('', '000000')

    /* =====================================================================
       2) คำนวณค่าใช้จ่ายระดับบรรทัด แล้วสรุปตามกลุ่ม
          Grouping : rec_type -> comm_rate -> product_group
       ===================================================================== */
    SELECT
        /* --- คีย์การจัดกลุ่มของรายงาน --- */
        l.rec_type,
        l.comm_rate,
        comm_rate_pct  = CONVERT(DECIMAL(9, 2), l.comm_rate * 100),
        group_caption  = CASE WHEN l.rec_type = 'P' THEN 'Policy ' ELSE 'Endorse ' END
                       + LTRIM(CONVERT(VARCHAR(20), CONVERT(DECIMAL(9, 2), l.comm_rate * 100)))
                       + '%',
        product_group  = l.product_group,

        /* --- CR ข้อ 10 : จำนวนกรมธรรม์ --- */
        policy_count   = COUNT(DISTINCT l.policy_key),

        /* --- CR ข้อ 11, 12 : เบี้ยสุทธิ / เบี้ยรวม --- */
        net_amount     = SUM(l.net),
        total_amount   = SUM(l.total),

        /* --- เบี้ยประกันภัย ปัดเศษทศนิยม --- */
        round_adj      = SUM(l.round_adj),

        /* --- CR ข้อ 13 - 15 : ค่าใช้จ่าย 1 + VAT 7% + TAX 3% --- */
        exp1           = SUM(l.exp1),
        exp1_vat7      = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.exp1 * 0.07, 2) ELSE 0 END),
        exp1_tax3      = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.exp1 * 0.03, 2) ELSE 0 END),

        /* --- CR ข้อ 16 - 18 : ค่าใช้จ่าย 2 + VAT 7% + TAX 3% --- */
        exp2           = SUM(l.exp2),
        exp2_vat7      = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.exp2 * 0.07, 2) ELSE 0 END),
        exp2_tax3      = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.exp2 * 0.03, 2) ELSE 0 END),

        /* --- CR ข้อ 19 - 21 : ค่าใช้พิมพ์กรมธรรม์ + VAT 7% + TAX 3% --- */
        print_fee      = SUM(l.print_fee),
        print_vat7     = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.print_fee * 0.07, 2) ELSE 0 END),
        print_tax3     = SUM(CASE WHEN @vat_registered = 1 THEN ROUND(l.print_fee * 0.03, 2) ELSE 0 END)
    FROM (
        SELECT
            x.rec_type,
            x.comm_rate,
            x.product_group,
            x.policy_key,
            x.net,
            x.total,
            x.round_adj,
            /* ค่าใช้จ่าย 1 : ส่วนที่ไม่เกินอัตรา คปภ. 18% */
            exp1 = ROUND(x.net * CASE WHEN x.comm_rate > 0.18
                                      THEN 0.18 ELSE x.comm_rate END, 2),
            /* ค่าใช้จ่าย 2 : ส่วนที่เกิน 18% */
            exp2 = ROUND(x.net * CASE WHEN x.comm_rate > 0.18
                                      THEN x.comm_rate - 0.18 ELSE 0 END, 2),
            /* ค่าใช้พิมพ์กรมธรรม์ 1% เฉพาะกรมธรรม์ที่ตัวแทนพิมพ์เอง */
            print_fee = ROUND(x.net * CASE WHEN x.print_epolicy IN ('Y', '1', 'Y/1')
                                           THEN 0.01 ELSE 0 END, 2)
        FROM #line x
    ) l
    GROUP BY l.rec_type, l.comm_rate, l.product_group
    /* เรียง Policy ก่อน Endorse, ภายในกลุ่มเรียง %ค่าใช้จ่ายจากน้อยไปมาก */
    ORDER BY l.rec_type ASC, l.comm_rate ASC, l.product_group ASC

    DROP TABLE #line
END
GO

GRANT EXECUTE ON dbo.proc_acvun7123_summary TO PUBLIC
GO
