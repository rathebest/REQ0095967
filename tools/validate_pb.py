#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
REQ0095968 - ตรวจสอบโครงสร้างไฟล์ PowerBuilder export ก่อนนำเข้า .pbl

ตรวจสิ่งที่ทำให้ Import ล้มเหลวหรือ Check Syntax ไม่ผ่านเป็นหลัก:
  .srd  - ชื่อใน $PBExportHeader$ ตรงกับชื่อไฟล์
        - ทุก column(...) อ้าง name= ที่ประกาศไว้ใน table(...)
        - id= ของ column(...) ตรงกับลำดับคอลัมน์ที่ประกาศใน table(...)
        - จำนวน argument ใน retrieve= ตรงกับ arguments=(...)
        - นิพจน์ใน band summary ต้องเป็น aggregate (PB ไม่ยอมให้อ้างคอลัมน์ตรง ๆ)
  .srw  - control ใน forward / global type / on create / on destroy ครบตรงกัน
        - ทุก prototype มี body

ใช้งาน:  python3 tools/validate_pb.py pb/*.srd pb/*.srw
คืนค่า exit code 1 เมื่อพบข้อผิดพลาด
"""
import io
import os
import re
import sys

AGG = re.compile(r'\b(sum|avg|count|first|last|max|min|median|mode|stdev|var|'
                 r'page|pageCount|today|now|cumulativeSum|cumulativePercent)\s*\(', re.I)


def check_srd(path, text):
    errs = []
    base = os.path.basename(path)
    name = base[:-4]

    m = re.match(r'\$PBExportHeader\$(\S+)', text)
    if not m:
        errs.append('ไม่พบบรรทัด $PBExportHeader$')
    elif m.group(1) != base:
        errs.append('$PBExportHeader$%s ไม่ตรงกับชื่อไฟล์ %s' % (m.group(1), base))

    if 'release 8;' not in text:
        errs.append("ไม่พบ 'release 8;'")

    # ---- คอลัมน์ที่ประกาศใน table(...) ตามลำดับ ----
    tm = re.search(r'^table\(.*?^\s*\)\s*$', text, re.S | re.M)
    if not tm:
        tm = re.search(r'^table\(.*?(?=^\w+\()', text, re.S | re.M)
    if not tm:
        errs.append('ไม่พบบล็อก table(...)')
        return errs
    tblock = tm.group(0)
    declared = []
    for chunk in tblock.split('column=(')[1:]:
        cm = re.search(r'\bname=(\w+)', chunk)
        if cm:
            declared.append(cm.group(1))
    if not declared:
        errs.append('บล็อก table(...) ไม่มีคอลัมน์')
        return errs
    index_of = {c: i + 1 for i, c in enumerate(declared)}

    # ---- retrieve= กับ arguments=(...) ----
    rm = re.search(r'retrieve="((?:[^"]|~")*)"', tblock)
    am = re.search(r'arguments=\(((?:[^()]|\([^()]*\))*)\)', tblock)
    if rm:
        nph = len(set(re.findall(r':(\w+)', rm.group(1))))
        nargs = len(re.findall(r'\("(\w+)"\s*,', am.group(1))) if am else 0
        if nph != nargs:
            errs.append('retrieve= มี placeholder %d ตัว แต่ arguments= ประกาศ %d ตัว' % (nph, nargs))

    # ---- column(band=...) ----
    for line in text.splitlines():
        if not line.startswith('column(band='):
            continue
        nm = re.search(r'\bname=(\w+)', line)
        im = re.search(r'\bid=(\d+)', line)
        if not nm:
            errs.append('column(...) ไม่มี name=  ->  %s' % line[:70])
            continue
        cname = nm.group(1)
        if cname not in index_of:
            errs.append('column name=%s ไม่ได้ประกาศใน table(...)' % cname)
            continue
        if im and int(im.group(1)) != index_of[cname]:
            errs.append('column name=%s ใช้ id=%s แต่ประกาศไว้ลำดับที่ %d'
                        % (cname, im.group(1), index_of[cname]))

    # ---- นิพจน์ใน band summary ต้องเป็น aggregate ----
    for line in text.splitlines():
        if 'band=summary' in line and 'expression="' in line:
            e = re.search(r'expression="((?:[^"]|~")*)"', line).group(1)
            if not AGG.search(e):
                errs.append('นิพจน์ใน band summary ไม่ได้ใช้ aggregate: %s' % e[:60])

    return errs


def check_srw(path, text):
    errs = []
    base = os.path.basename(path)

    m = re.match(r'\$PBExportHeader\$(\S+)', text)
    if not m:
        errs.append('ไม่พบบรรทัด $PBExportHeader$')
    elif m.group(1) != base:
        errs.append('$PBExportHeader$%s ไม่ตรงกับชื่อไฟล์ %s' % (m.group(1), base))

    win = base[:-4]

    fm = re.search(r'^forward\s*$(.*?)^end forward\s*$', text, re.S | re.M)
    if not fm:
        errs.append('ไม่พบบล็อก forward ... end forward')
        return errs
    fwd = set(re.findall(r'^type (\w+) from ', fm.group(1), re.M))

    after = text[fm.end():]
    gm = re.search(r'^global type %s from window\s*$(.*?)^end type\s*$' % re.escape(win),
                   after, re.S | re.M)
    if not gm:
        errs.append('ไม่พบ global type %s from window' % win)
        return errs
    gdecl = set(re.findall(r'^(\w+) \1\s*$', gm.group(1), re.M))

    created = set(re.findall(r'this\.(\w+) = create \1', text))
    destroyed = set(re.findall(r'destroy\(this\.(\w+)\)', text))

    for label, s in (('global type', gdecl), ('on create', created), ('on destroy', destroyed)):
        miss = fwd - s
        if miss:
            errs.append('control ประกาศใน forward แต่ไม่พบใน %s: %s' % (label, ', '.join(sorted(miss))))
    extra = created - fwd
    if extra:
        errs.append('control ถูก create แต่ไม่ได้ประกาศใน forward: %s' % ', '.join(sorted(extra)))

    # ---- prototypes กับ body ----
    pm = re.search(r'^forward prototypes\s*$(.*?)^end prototypes\s*$', text, re.S | re.M)
    if pm:
        protos = set(re.findall(r'function \w+(?:\{\d+\})? (\w+) ?\(', pm.group(1)))
        bodies = set(re.findall(r'^public function \w+(?:\{\d+\})? (\w+) ?\(.*?\);', text, re.M))
        miss = protos - bodies
        if miss:
            errs.append('prototype ไม่มี body: %s' % ', '.join(sorted(miss)))
        extra = bodies - protos
        if extra:
            errs.append('function มี body แต่ไม่ได้ประกาศ prototype: %s' % ', '.join(sorted(extra)))

    return errs


def main(argv):
    files = argv[1:]
    if not files:
        print('usage: validate_pb.py <files...>')
        return 2
    total = 0
    for path in files:
        text = io.open(path, encoding='utf-8').read()
        errs = check_srd(path, text) if path.endswith('.srd') else check_srw(path, text)
        if errs:
            total += len(errs)
            print('FAIL %s' % path)
            for e in errs:
                print('   - %s' % e)
        else:
            print('OK   %s' % path)
    print('\n%d error(s)' % total)
    return 1 if total else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
