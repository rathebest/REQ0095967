/* =============================================================================
   CR         : REQ0095968 - สรุปยอดรายงานตัดชำระเบี้ยตามใบนำส่งเงินชำระเบี้ย
   Object     : [dbo].[fn_acvun_debitno_key]  (Scalar Function - NEW)
   Purpose    : แปลงเลขที่ใบนำส่งเงินชำระเบี้ยจากรูปแบบที่ผู้ใช้กรอก (NNNNN/YY)
                ให้เป็นรูปแบบที่จัดเก็บในตาราง debit_list.debit_no = CHAR(7) = YYNNNNN
                เพื่อให้เปรียบเทียบช่วง From - To ด้วย BETWEEN ได้ถูกต้อง
                (คีย์ YYNNNNN เรียงตามปีแล้วตามเลขที่วิ่ง จึงเทียบช่วงได้ตรงตามความหมาย)
   Target DB  : ฐานข้อมูลระบบงานภาคสมัครใจ (VMI) ที่หน้าจอ w_acvun7123 เชื่อมต่อ
   Created    : REQ0095968
   ---------------------------------------------------------------------------
   ตัวอย่าง
     SELECT dbo.fn_acvun_debitno_key('00025/68')  -->  '6800025'
     SELECT dbo.fn_acvun_debitno_key('25/68')     -->  '6800025'
     SELECT dbo.fn_acvun_debitno_key('6800025')   -->  '6800025'  (ส่งรูปแบบจัดเก็บมาตรง ๆ)
     SELECT dbo.fn_acvun_debitno_key('')          -->  NULL
     SELECT dbo.fn_acvun_debitno_key('AB/68')     -->  NULL       (ไม่ใช่ตัวเลข)
   ============================================================================= */
IF OBJECT_ID('dbo.fn_acvun_debitno_key') IS NOT NULL
    DROP FUNCTION dbo.fn_acvun_debitno_key
GO

CREATE FUNCTION [dbo].[fn_acvun_debitno_key]
(
    @disp VARCHAR(20)          -- เลขที่ใบนำส่งฯ รูปแบบ 'NNNNN/YY' หรือ 'YYNNNNN'
)
RETURNS CHAR(7)
AS
BEGIN
    DECLARE @s   VARCHAR(20),
            @run VARCHAR(20),
            @yy  VARCHAR(20),
            @p   INT

    SET @s = LTRIM(RTRIM(ISNULL(@disp, '')))

    IF @s = ''
        RETURN NULL

    SET @p = CHARINDEX('/', @s)

    /* ---- ไม่มีตัวคั่น : ถือว่าเป็นรูปแบบจัดเก็บ YYNNNNN อยู่แล้ว ---- */
    IF @p = 0
    BEGIN
        IF LEN(@s) = 7 AND @s NOT LIKE '%[^0-9]%'
            RETURN @s
        RETURN NULL
    END

    SET @run = LTRIM(RTRIM(LEFT(@s, @p - 1)))          -- เลขที่วิ่ง
    SET @yy  = LTRIM(RTRIM(SUBSTRING(@s, @p + 1, 20))) -- ปี พ.ศ. 2 หลักท้าย

    /* ---- ต้องเป็นตัวเลขล้วนและไม่ว่าง ---- */
    IF @run = '' OR @yy = ''
        RETURN NULL
    IF @run LIKE '%[^0-9]%' OR @yy LIKE '%[^0-9]%'
        RETURN NULL
    IF LEN(@run) > 5 OR LEN(@yy) > 2
        RETURN NULL

    SET @run = RIGHT('00000' + @run, 5)
    SET @yy  = RIGHT('00' + @yy, 2)

    RETURN @yy + @run
END
GO

GRANT EXECUTE ON dbo.fn_acvun_debitno_key TO PUBLIC
GO
