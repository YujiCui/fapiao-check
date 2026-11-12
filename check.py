#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""发票体检：对发票清单做勾稽和风险筛查，输出异常清单与摘要。

    python check.py 清单.xlsx --side in --name 企业全称 --taxid 税号 --start 2026-07-01 --end 2026-09-30

只用标准库；读 .xlsx 需要 openpyxl。控制台只打印一行计数，结果看输出文件。
"""
import argparse
import csv
import datetime as dt
import io
import os
import re
import sys
from collections import defaultdict
from decimal import Decimal, InvalidOperation

# ================= 口径：政策或习惯变了改这里 =================
RATES_NOW = {Decimal(x) for x in "0 0.005 0.01 0.015 0.02 0.03 0.05 0.06 0.09 0.13".split()}
RATES_OLD = {Decimal(x) for x in "0.10 0.11 0.16 0.17".split()}
SUM_TOL = Decimal("0.01")      # 金额+税额 与 价税合计 的容差
TAX_TOL = Decimal("1.27")      # 金额×税率 与 税额 的容差（只影响小额差异是否提示）
ROUND_MIN = Decimal("50000")   # R8：整票价税合计不低于此数，且是 ROUND_UNIT 的整数倍
ROUND_UNIT = Decimal("10000")
SENS_MIN = Decimal("10000")    # R9：业务真实性类品目的起报金额
KW_REAL = ["咨询", "顾问", "服务费", "会议", "会务", "推广", "宣传", "劳务", "办公用品",
           "礼品", "培训", "技术服务", "信息服务", "佣金", "手续费"]
KW_DEDUCT = ["餐饮", "餐费", "娱乐", "贷款", "利息", "福利", "美容", "洗浴", "家政"]
BAD_STATUS = ["作废", "红冲", "异常", "失控"]
OK_RISK = {"", "正常", "无", "-", "--"}
FREE_RATE = ("免税", "不征税")

ALIASES = {  # 标准名 -> 认得的表头，靠前的优先
    "数电票号码": ["数电票号码", "数电发票号码", "全电发票号码"],
    "发票代码": ["发票代码"],
    "发票号码": ["发票号码", "票号", "发票号"],
    "开票日期": ["开票日期", "开票时间", "日期"],
    "票种": ["发票票种", "票种", "发票类型", "发票种类"],
    "状态": ["发票状态", "状态"],
    "风险等级": ["发票风险等级", "风险等级"],
    "销方名称": ["销方名称", "销售方名称", "销方", "开票方名称"],
    "销方税号": ["销方识别号", "销售方识别号", "销方税号", "销售方纳税人识别号", "销方纳税人识别号"],
    "购方名称": ["购买方名称", "购方名称", "购方", "受票方名称"],
    "购方税号": ["购方识别号", "购买方识别号", "购方税号", "购买方纳税人识别号", "购方纳税人识别号"],
    "品目": ["货物或应税劳务名称", "货物或应税劳务、服务名称", "项目名称", "品目", "品名", "商品名称"],
    "金额": ["金额", "不含税金额", "合计金额"],
    "税率": ["税率", "税率/征收率", "征收率"],
    "税额": ["税额", "合计税额"],
    "价税合计": ["价税合计", "含税金额", "价税合计金额"],
    "备注": ["备注"],
}

RULES = [
    ("R1", "勾稽：金额+税额=价税合计，金额×税率=税额"),
    ("R2", "重复发票"),
    ("R3", "发票状态或风险等级异常"),
    ("R4", "抬头或税号与本企业不符"),
    ("R5", "税号位数或校验位异常"),
    ("R6", "税率不在常见集合"),
    ("R7", "开票日期为空、无法识别或不在检查期间"),
    ("R8", "大额整数发票"),
    ("R9", "敏感品目"),
]

USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_W = (1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28)
ID_W = (7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2)


def die(msg):
    sys.stderr.write("ERROR: %s\n" % msg)
    sys.exit(2)


def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def money(v):
    if v is None or isinstance(v, bool):
        return None
    if not isinstance(v, (int, float, Decimal)):
        v = re.sub(r"[,，\s¥￥元]", "", str(v))
        if v.startswith("(") and v.endswith(")"):
            v = "-" + v[1:-1]
    try:
        return Decimal(str(v)).quantize(Decimal("0.01"))
    except InvalidOperation:
        return None


def parse_rate(v):
    """返回 (税率, 文本)。文本非空 = 免税/不征税，或认不出的原文。"""
    if v is None or isinstance(v, bool):
        return None, ""
    pct = False
    if isinstance(v, (int, float, Decimal)):
        d = Decimal(str(v))
    else:
        t = str(v).strip()
        if t in ("", "-", "--", "*", "***"):
            return None, ""
        pct = t[-1] in "%％"
        try:
            d = Decimal(t.rstrip("%％").strip())
        except InvalidOperation:
            return None, t
    if pct or d >= 1:
        d = d / 100
    return d.quantize(Decimal("0.0001")), ""


def parse_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    t = clean(v)
    m = (re.match(r"(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})", t)
         or re.match(r"(\d{4})(\d{2})(\d{2})(?:\d{6})?$", t))
    try:
        if m:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if re.fullmatch(r"\d{5}", t) and 30000 < int(t) < 60000:  # Excel 日期序号
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(t))
    except ValueError:
        pass
    return None


def uscc_check_char(body17):
    s = sum(USCC_CHARS.index(ch) * w for ch, w in zip(body17, USCC_W))
    return USCC_CHARS[(31 - s % 31) % 31]


def taxid_issue(c):
    """税号没问题返回 None，否则返回问题描述。15/17/20 位旧税号和打码的不查。"""
    c = clean(c).upper()
    if not c or "*" in c or len(c) in (15, 17, 20):
        return None
    if len(c) != 18:
        return "位数为 %d" % len(c)
    if all(ch in USCC_CHARS for ch in c) and uscc_check_char(c[:17]) == c[17]:
        return None
    if re.fullmatch(r"\d{17}[\dX]", c):  # 自然人：身份证号
        if "10X98765432"[sum(int(ch) * w for ch, w in zip(c, ID_W)) % 11] == c[17]:
            return None
    return "18 位，校验位不对"


def norm_name(s):
    return re.sub(r"\s+", "", s).replace("（", "(").replace("）", ")").upper()


def norm_header(v):
    return re.sub(r"[（(].*?[)）]|\s|﻿", "", clean(v))


def find_header(rows):
    """前 15 行里命中别名最多的一行当表头。返回 (行下标, {标准名: 列下标})。"""
    best_i, best = -1, {}
    for i, row in enumerate(rows[:15]):
        names = [norm_header(c) for c in row]
        cols = {}
        for std, alts in ALIASES.items():
            for alt in alts:
                if alt in names:
                    cols[std] = names.index(alt)
                    break
        if len(cols) > len(best):
            best_i, best = i, cols
    return best_i, best


def read_table(path, sheet):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            die("reading .xlsx needs openpyxl (pip install openpyxl), or save the sheet as CSV")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if sheet and sheet not in wb.sheetnames:
            die("sheet not found: %s" % sheet)
        best = None
        for name in ([sheet] if sheet else wb.sheetnames):  # 多个工作表时取列最全的
            ws = wb[name]
            ws.reset_dimensions()
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            i, cols = find_header(rows)
            if best is None or len(cols) > len(best[3]):
                best = (name, rows, i, cols)
        return best
    if ext == ".xls":
        die(".xls is not supported: save as .xlsx or CSV")
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "gb18030"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        die("cannot decode the file as UTF-8 or GBK")
    first = text.split("\n", 1)[0]
    delim = "\t" if first.count("\t") > first.count(",") else ","
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delim))
    i, cols = find_header(rows)
    return "", rows, i, cols


def fmt(x):
    return "{:,.2f}".format(x or Decimal(0))


def pct(d):
    return "{:f}%".format((d * 100).normalize())


def main():
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, "reconfigure"):
            s.reconfigure(errors="backslashreplace")
    ap = argparse.ArgumentParser(description="Invoice list checker. See SKILL.md.")
    ap.add_argument("path", help="invoice list (.xlsx or .csv)")
    ap.add_argument("--sheet", help="worksheet name; default: the one with most recognized columns")
    ap.add_argument("--side", choices=["in", "out"], default="in", help="in = purchases, out = sales")
    ap.add_argument("--name", help="the audited company's full name")
    ap.add_argument("--taxid", help="the audited company's tax id")
    ap.add_argument("--start", help="period start, YYYY-MM-DD")
    ap.add_argument("--end", help="period end, YYYY-MM-DD")
    ap.add_argument("--per-invoice", action="store_true",
                    help="list has one row per invoice: report every repeated invoice number")
    ap.add_argument("--out", help="output directory; default: next to the list")
    a = ap.parse_args()

    start, end = parse_date(a.start), parse_date(a.end)
    if (a.start and not start) or (a.end and not end):
        die("--start/--end must look like 2026-07-01")
    if not os.path.isfile(a.path):
        die("file not found: %s" % a.path)

    sheet, rows, hdr, cols = read_table(a.path, a.sheet)
    if len(cols) < 3:
        die("header row not recognized; rename the columns to the standard names in SKILL.md")
    mine, other = ("购方", "销方") if a.side == "in" else ("销方", "购方")
    mine_k, other_k = ("buyer", "seller") if a.side == "in" else ("seller", "buyer")
    has_id = any(c in cols for c in ("数电票号码", "发票号码", "开票日期"))

    def cell(row, std):
        j = cols.get(std)
        return row[j] if j is not None and j < len(row) else None

    recs, skipped, sci = [], 0, 0
    for n, row in enumerate(rows[hdr + 1:], start=hdr + 2):  # n = 清单里的行号
        if not any(clean(c) for c in row):
            continue
        key = clean(cell(row, "数电票号码")) or clean(cell(row, "发票代码")) + clean(cell(row, "发票号码"))
        date = parse_date(cell(row, "开票日期"))
        if has_id and date is None and not re.search(r"\d", key):  # 合计行、说明行
            skipped += 1
            continue
        if re.fullmatch(r"\d(\.\d+)?[eE]\+?\d+", key):  # 票号被 Excel 改成科学计数法
            sci += 1
            key = ""
        rate, rate_txt = parse_rate(cell(row, "税率"))
        recs.append(dict(
            row=n, key=key, gkey=key or "#%d" % n, date=date, date_raw=clean(cell(row, "开票日期")),
            kind=clean(cell(row, "票种")), status=clean(cell(row, "状态")), risk=clean(cell(row, "风险等级")),
            seller=clean(cell(row, "销方名称")), seller_id=clean(cell(row, "销方税号")),
            buyer=clean(cell(row, "购方名称")), buyer_id=clean(cell(row, "购方税号")),
            item=clean(cell(row, "品目")), amt=money(cell(row, "金额")), tax=money(cell(row, "税额")),
            total=money(cell(row, "价税合计")), rate=rate, rate_txt=rate_txt, memo=clean(cell(row, "备注")),
        ))
    if not recs:
        die("no data rows found")

    # 按票号归组。whole=True 表示明细行里重复写的是整票合计。
    groups = defaultdict(list)
    for r in recs:
        groups[r["gkey"]].append(r)
    ginfo = {}
    for k, g in groups.items():
        parts = sum((x["amt"] or 0) + (x["tax"] or 0) for x in g)
        totals = [x["total"] for x in g if x["total"] is not None]
        whole = (len(g) > 1 and len(totals) == len(g) and len(set(totals)) == 1
                 and abs(parts - totals[0]) <= SUM_TOL * len(g))
        ginfo[k] = (whole, totals[0] if whole else sum(totals) if len(totals) == len(g) else parts)

    def amount(r):
        if r["total"] is not None and not ginfo[r["gkey"]][0]:
            return r["total"]
        if r["amt"] is not None:
            return r["amt"] + (r["tax"] or 0)
        return None

    hits = []

    def hit(r, rule, msg, amt=None):
        hits.append((r, rule, msg, amount(r) if amt is None else amt))

    for r in recs:
        amt, tax, total = r["amt"], r["tax"], r["total"]
        if None not in (amt, tax, total) and abs(amt + tax - total) > SUM_TOL and not ginfo[r["gkey"]][0]:
            hit(r, "R1", "金额+税额=%s，价税合计=%s，差 %s" % (fmt(amt + tax), fmt(total), fmt(amt + tax - total)))
        if None not in (amt, tax, r["rate"]) and "差额" not in r["memo"]:
            want = (amt * r["rate"]).quantize(Decimal("0.01"))
            if abs(want - tax) > TAX_TOL:
                hit(r, "R1", "金额×税率=%s，税额=%s，差 %s" % (fmt(want), fmt(tax), fmt(want - tax)))

        if any(b in r["status"] for b in BAD_STATUS):
            hit(r, "R3", "发票状态：%s" % r["status"])
        if r["risk"] not in OK_RISK:
            hit(r, "R3", "风险等级：%s" % r["risk"])

        nm, tid = r[mine_k], r[mine_k + "_id"]
        if a.name and nm and norm_name(nm) != norm_name(a.name):
            hit(r, "R4", "%s名称为「%s」" % (mine, nm))
        if a.taxid and tid and "*" not in tid and tid.upper() != a.taxid.strip().upper():
            hit(r, "R4", "%s税号为 %s" % (mine, tid))

        for who, k in (("销方", "seller_id"), ("购方", "buyer_id")):
            problem = taxid_issue(r[k])
            if problem:
                hit(r, "R5", "%s税号 %s：%s" % (who, r[k], problem))

        if r["rate_txt"] and not any(f in r["rate_txt"] for f in FREE_RATE):
            hit(r, "R6", "税率无法识别：%s" % r["rate_txt"])
        elif r["rate"] is not None and r["rate"] in RATES_OLD:
            hit(r, "R6", "税率 %s 是历史税率，核对开票日期" % pct(r["rate"]))
        elif r["rate"] is not None and r["rate"] not in RATES_NOW:
            hit(r, "R6", "税率 %s 不在常见集合" % pct(r["rate"]))

        if "开票日期" in cols:
            if r["date"] is None:
                hit(r, "R7", "开票日期" + ("无法识别：%s" % r["date_raw"] if r["date_raw"] else "为空"))
            elif (start and r["date"] < start) or (end and r["date"] > end):
                hit(r, "R7", "开票日期 %s 不在检查期间" % r["date"])

        if r["item"]:
            val = amount(r)
            kw = [k for k in KW_REAL if k in r["item"]]
            if kw and val is not None and val >= SENS_MIN:
                hit(r, "R9", "品目含「%s」，核对合同、成果和付款" % "、".join(kw))
            kw = [k for k in KW_DEDUCT if k in r["item"]]
            if kw and ("专" in r["kind"] or not r["kind"]) and (val is None or val > 0):
                hit(r, "R9", "品目含「%s」，核对抵扣处理" % "、".join(kw))

    for k, g in groups.items():
        if not k.startswith("#"):
            first = {}
            for x in g:  # 明细清单里同票号多行正常，只报内容完全相同的行
                sig = None if a.per_invoice else (x["item"], x["amt"], x["rate"], x["tax"], x["total"])
                if sig in first:
                    hit(x, "R2", "与第 %d 行票号%s相同" % (first[sig], "" if a.per_invoice else "和内容"))
                else:
                    first[sig] = x["row"]
        inv = ginfo[k][1]
        if inv and inv >= ROUND_MIN and inv % ROUND_UNIT == 0:
            hit(g[0], "R8", "整票价税合计 %s，是 %s 的整数倍" % (fmt(inv), fmt(ROUND_UNIT)), inv)

    # ---------------- 输出 ----------------
    stem = os.path.splitext(os.path.basename(a.path))[0]
    out_dir = a.out or os.path.dirname(os.path.abspath(a.path))
    os.makedirs(out_dir, exist_ok=True)
    f_csv = os.path.join(out_dir, stem + "_异常清单.csv")
    f_md = os.path.join(out_dir, stem + "_摘要.md")

    hits.sort(key=lambda h: (h[1], -(h[3] or 0), h[0]["row"]))
    with open(f_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["行号", "规则", "说明", "票号", "开票日期", other + "名称", "品目",
                    "金额", "税额", "价税合计", "复核意见"])
        for r, rule, msg, _ in hits:
            key = "\t" + r["key"] if r["key"].isdigit() and len(r["key"]) > 11 else r["key"]  # 防 Excel 转科学计数法
            w.writerow([r["row"], rule, msg, key, r["date"] or r["date_raw"], r[other_k], r["item"],
                        r["amt"], r["tax"], r["total"], ""])

    def has(*names):
        return all(n in cols for n in names)

    not_run = []
    if not (has("金额", "税额") and (has("价税合计") or has("税率"))):
        not_run.append(("R1", "缺金额、税额，或缺价税合计和税率"))
    if not (has("数电票号码") or has("发票号码")):
        not_run.append(("R2", "缺票号列"))
    if not (has("状态") or has("风险等级")):
        not_run.append(("R3", "缺发票状态、风险等级列"))
    if not (a.name or a.taxid):
        not_run.append(("R4", "没给本企业名称或税号"))
    elif not (has(mine + "名称") or has(mine + "税号")):
        not_run.append(("R4", "缺%s名称、税号列" % mine))
    if not (has("销方税号") or has("购方税号")):
        not_run.append(("R5", "缺税号列"))
    if not has("税率"):
        not_run.append(("R6", "缺税率列"))
    if not has("开票日期"):
        not_run.append(("R7", "缺开票日期列"))
    elif not (start or end):
        not_run.append(("R7", "没给检查期间，只查了日期能否识别"))
    if not (has("价税合计") or has("金额")):
        not_run.append(("R8", "缺金额列"))
    if not has("品目"):
        not_run.append(("R9", "缺品目列"))

    by = defaultdict(lambda: [0, Decimal(0)])
    for k, g in groups.items():
        by[g[0][other_k] or "（空）"][0] += 1
        by[g[0][other_k] or "（空）"][1] += ginfo[k][1] or 0
    grand = sum(v[1] for v in by.values())
    top = sorted(by.items(), key=lambda kv: -kv[1][1])[:5]
    hit_inv = {r["gkey"] for r, _, _, _ in hits}
    n_red = sum(1 for r in recs if (amount(r) or 0) < 0)

    L = ["# 发票体检摘要", ""]
    L.append("- 清单：%s%s" % (os.path.basename(a.path), "（工作表 %s）" % sheet if sheet else ""))
    L.append("- 口径：%s；本企业 %s %s；检查期间 %s 至 %s" % (
        "进项" if a.side == "in" else "销项", a.name or "未提供", a.taxid or "",
        start or "未提供", end or "未提供"))
    L.append("- %d 行，%d 张发票，红字 %d 行；跳过 %d 行（无票号且无日期）" % (
        len(recs), len(groups), n_red, skipped))
    L.append("- 金额 %s，税额 %s，价税合计 %s" % (
        fmt(sum(r["amt"] or 0 for r in recs)), fmt(sum(r["tax"] or 0 for r in recs)), fmt(grand)))
    L.append("- 命中 %d 行次，涉及 %d 张发票" % (len(hits), len(hit_inv)))
    L += ["", "## 规则命中", "", "| 规则 | 查什么 | 行次 | 涉及价税合计 |", "|---|---|---|---|"]
    for code, desc in RULES:
        seen, total, cnt = set(), Decimal(0), 0
        for r, rule, _, val in hits:
            if rule == code:
                cnt += 1
                if r["row"] not in seen:
                    seen.add(r["row"])
                    total += val or 0
        L.append("| %s | %s | %d | %s |" % (code, desc, cnt, fmt(total) if cnt else "-"))
    L += ["", "## 没跑或没跑全的规则", ""]
    L += ["- %s：%s" % x for x in not_run] or ["- 无"]
    if sci:
        L += ["", "## 数据质量", "",
              "- %d 行的票号是科学计数法（被 Excel 改写过），这些行没法查重。用原始导出文件重跑。" % sci]
    L += ["", "## 前五大%s" % other, "", "| 名称 | 发票数 | 价税合计 | 占比 |", "|---|---|---|---|"]
    for name, (cnt, val) in top:
        L.append("| %s | %d | %s | %s |" % (name, cnt, fmt(val), "{:.1%}".format(val / grand) if grand > 0 else "-"))
    L += ["", "## 脚本没查的", "",
          "- %s是否非正常户、走逃户" % other,
          "- 业务是否真实、资金是否回流",
          "- 品目与主营业务是否相称（R8、R9 命中的行需要人工或模型复核）", ""]
    with open(f_md, "w", encoding="utf-8") as f:
        f.write("\n".join(L))

    print("OK rows=%d invoices=%d hits=%d not_run=%s" % (
        len(recs), len(groups), len(hits), ",".join(sorted({x[0] for x in not_run})) or "-"))
    print(f_csv)
    print(f_md)


if __name__ == "__main__":
    main()
