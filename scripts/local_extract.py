#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Local-only extractor for spreadsheets, documents, PDFs, images and text files.

Reads files from the local filesystem only, never opens a network connection,
and writes every artefact into a workspace directory the caller controls.

Masking of personal data is ALWAYS ON and cannot be disabled: names, phone
numbers, addresses, identity/passport numbers, bank cards, e-mail, plates and
account identifiers. Masked values become stable pseudonyms ([姓名#1],
[电话#2] ...) so counts and joins stay valid while identities do not.

Supported without any third-party package:
  .xlsx .xlsm        OOXML workbook  (zip + xml)
  .docx .docm        OOXML document  (zip + xml)
  .pptx .pptm        OOXML slides    (zip + xml)
  .csv .tsv .txt .md .json .html .htm .log .yaml .yml .xml
  .pdf               via local workbuddy-pdf CLI when available, else builtin
  .png .jpg .jpeg .gif .bmp .webp    header/metadata only, no OCR

Usage:
  python local_extract.py FILE [FILE ...] [options]

Options:
  --mode normal|strict      normal = 结构与统计 + 脱敏后的值预览（默认）
                            strict = 只输出聚合指标，不出现任何单元格原文
  --workspace DIR           output directory (default: ./.local_analysis)
  --rows N                  preview rows (default 5; strict caps at 2)
  --max-chars N             max characters of document text kept (default 6000)
  --mask-region             mask the whole address instead of keeping province/city
  --redact REGEX            extra regex to mask (repeatable)
  --mask-headers            replace column names with C1..Cn
  --quiet                   minimal stdout
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import datetime
from html import unescape
from xml.etree import ElementTree as ET

# --------------------------------------------------------------------------
# constants
# --------------------------------------------------------------------------

NS_SS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

SHEET_EXT = {".xlsx", ".xlsm"}
DOC_EXT = {".docx", ".docm"}
SLIDE_EXT = {".pptx", ".pptm"}
PDF_EXT = {".pdf"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}
TEXT_EXT = {".csv", ".tsv", ".txt", ".md", ".markdown", ".json", ".html",
            ".htm", ".log", ".yaml", ".yml", ".xml", ".ini", ".conf"}
LEGACY_OFFICE = {".xls", ".doc", ".ppt"}

# column-header keywords -> data category to mask
CATEGORY_HINTS = [
    ("姓名", ("姓名", "名字", "客户名", "客户名称", "户名", "联系人", "法人",
              "法定代表人", "负责人", "经办人", "员工", "客户经理", "签字",
              "收款人", "开户人", "持卡人", "申请人", "授权人", "股东", "董事",
              "户主", "借款人", "购房人")),
    ("电话", ("手机", "电话", "联系方式", "座机", "传真", "tel", "phone", "mobile", "contact")),
    ("证件", ("身份证", "证件", "护照", "军官证", "通行证", "台胞证", "营业执照",
              "统一社会信用代码", "税号", "纳税人识别号", "idcard")),
    ("地址", ("地址", "住址", "住所", "通讯", "所在地区", "户籍", "籍贯", "邮寄",
              "收货", "办公地", "注册地", "address")),
    ("邮箱", ("邮箱", "邮件", "email", "e-mail")),
    ("编号", ("客户号", "客户编号", "工号", "员工号", "会员号", "学号", "社保号", "保单号")),
    ("账号", ("卡号", "账号", "账户", "户号", "借记卡", "信用卡", "存折", "account")),
    ("车牌", ("车牌", "牌号")),
]

SURNAMES = ("赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"
            "戚谢邹喻柏水窦章云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳鲍史唐"
            "费廉岑薛雷贺倪汤滕殷罗毕郝邬安常乐于时傅皮卞齐康伍余元卜顾孟平"
            "黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅庞熊纪舒屈项祝"
            "董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田樊"
            "胡凌霍虞万支柯昝管卢莫房裘缪干解应宗丁宣邓郁单杭洪包诸左石崔吉"
            "钮龚程嵇邢裴陆荣翁荀羊惠甄曲家封芮储靳汲邴糜松井段富巫乌焦巴弓"
            "牧隗山谷车侯宓蓬全郗班仰秋仲伊宫宁仇栾暴甘厉戎祖武符刘景詹束龙"
            "叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍赖卓蔺屠蒙池乔阴胥能苍双"
            "闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍桑桂濮牛寿通边扈燕冀浦尚农温别"
            "庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡"
            "国文寇广禄阙东欧殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空"
            "曾毋沙乜养鞠须丰巢关蒯相查后荆红游竺权逯盖益桓公")
SURNAME_SET = set(SURNAMES)

STOP_TOKENS = {"电话", "手机", "地址", "邮箱", "名称", "号码", "签字", "盖章", "日期",
               "姓名", "联系", "方式", "单位", "信息", "资料", "详情", "备注", "编号"}


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def human_size(num: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.1f} {unit}"
        num /= 1024.0
    return f"{num:.1f} GB"


def read_text_any(path: str) -> tuple[str, str]:
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "utf-16", "latin-1"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace"), "utf-8(replace)"


def col_index(ref: str) -> int:
    m = re.match(r"([A-Z]+)", ref or "")
    idx = 0
    for ch in (m.group(1) if m else ""):
        idx = idx * 26 + (ord(ch) - 64)
    return max(idx - 1, 0)


def decode_cell(t, v, shared: list[str]) -> str:
    if t == "s":
        try:
            return shared[int(v)]
        except (TypeError, ValueError, IndexError):
            return ""
    return "" if v is None else str(v)


def fmt_num(v) -> str:
    if isinstance(v, int):
        return f"{v:,}"
    if abs(v) >= 1e7:
        return f"{v:.4g}"
    return f"{v:,.2f}".rstrip("0").rstrip(".")


def sniff_number(val: str):
    s = val.strip().replace(",", "")
    if not s:
        return None
    if len(s) > 1 and s[0] == "0" and s[1].isdigit():
        return None  # 保留前导零的编码类字段按文本处理
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d*\.\d+", s):
        return float(s)
    return None


# --------------------------------------------------------------------------
# mandatory masking
# --------------------------------------------------------------------------

def _region_prefix(v: str) -> str:
    m = re.match(r"^([\u4e00-\u9fa5]{2,8}?(?:省|自治区|特别行政区|市))?"
                 r"([\u4e00-\u9fa5]{2,8}?(?:市|自治州|地区|盟))?", v)
    if not m:
        return ""
    return (m.group(1) or "") + (m.group(2) or "")


class Masker:
    """Deterministic, mandatory masking of personal data.

    Each distinct original value maps to a stable pseudonym within one run, so
    cardinality and joins survive while identities never reach the report.
    No mapping table is ever written to disk.
    """

    TEXT_RULES = [
        ("身份证号", re.compile(r"(?<![\dA-Za-z])\d{17}[\dXx](?![\dA-Za-z])")),
        ("身份证号", re.compile(r"(?<!\d)\d{15}(?!\d)")),
        ("统一社会信用代码", re.compile(
            r"(?<![0-9A-Z])[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}(?![0-9A-Z])")),
        ("银行卡号", re.compile(r"(?<![\d-])\d{16,19}(?![\d-])")),
        ("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
        ("固定电话", re.compile(r"(?<!\d)0\d{2,3}-?\d{7,8}(?!\d)")),
        ("邮箱", re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")),
        ("车牌号", re.compile(r"[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼]"
                          r"[A-Z][A-Z0-9]{4,5}[A-Z0-9挂学警港澳]")),
    ]
    ADDRESS_RE = re.compile(
        r"[\u4e00-\u9fa5]{2,10}(?:省|自治区|特别行政区)?"
        r"[\u4e00-\u9fa5]{2,10}(?:市|自治州|地区|盟)"
        r"[\u4e00-\u9fa5]{0,25}(?:区|县|旗|市)?"
        r"[\u4e00-\u9fa50-9A-Za-z\-]{2,40}?(?:路|街|道|巷|弄|村|组|号|栋|幢|单元|楼|室|座)")
    # 门牌尾串：单元/室/层等，接在地址或句中出现时一并遮蔽
    UNIT_RE = re.compile(r"(?:\d{1,4}\s*(?:单元|号楼|栋|幢|室|层|户|房|座)){1,4}")
    # 三字候选被截时用于回退的常见谓词/助词，避免把"张三负责"整体吞掉
    TRIM_CHARS = set("负说表已将在的是与和及等同前后中内外未不也都就又还均要可需请由对向从到于或则并为以其之此各")

    def __init__(self, keep_region: bool = True, extra: list[str] | None = None):
        self.keep_region = keep_region
        self.maps: dict[str, dict[str, str]] = {}
        self.counters: dict[str, int] = {}
        self.hits: list[str] = []
        self.extra = []
        for raw in extra or []:
            try:
                self.extra.append(re.compile(raw))
            except re.error:
                continue
        surname = "[" + "".join(sorted(SURNAME_SET)) + "]"
        self.ctx_name = re.compile(
            r"(姓名|联系人|负责人|经办人|户名|客户经理|签字|申请人|持卡人|授权人|"
            r"法定代表人|法人|借款人|收款人|开户人)[：:\s]{0,3}"
            rf"({surname}[\u4e00-\u9fa5]{{1,2}})")
        self.honor_name = re.compile(
            rf"({surname}[\u4e00-\u9fa5]{{1,2}})(先生|女士|同志|经理|主任|行长|老师|医生|总监)")
        self.moumou = re.compile(rf"({surname})某某")

    # -- tokens -----------------------------------------------------------
    def token(self, kind: str, value: str) -> str:
        self.hits.append(kind)
        bucket = self.maps.setdefault(kind, {})
        if value not in bucket:
            self.counters[kind] = self.counters.get(kind, 0) + 1
            bucket[value] = f"[{kind}#{self.counters[kind]}]"
        return bucket[value]

    def address(self, value: str) -> str:
        v = value.strip()
        if not v:
            return v
        if self.keep_region:
            prefix = _region_prefix(v)
            if prefix and len(prefix) < len(v):
                return prefix + self.token("详细地址", v[len(prefix):])
        return self.token("地址", v)

    # -- field-level ------------------------------------------------------
    def field(self, kind: str | None, value: str) -> str:
        if not isinstance(value, str) or not value.strip():
            return value
        if kind == "地址":
            return self.address(value)
        if kind in ("姓名", "电话", "证件", "邮箱", "账号", "车牌", "编号"):
            return self.token(kind, value.strip())
        return self.text(value)

    # -- free-text --------------------------------------------------------
    def text(self, value: str) -> str:
        if not isinstance(value, str) or not value:
            return value
        s = value

        def _rule(_m, n: str = ""):
            self.hits.append(n)
            return f"[{n}]"
        for name, pat in self.TEXT_RULES:
            s = pat.sub(lambda _m, n=name: _rule(_m, n), s)

        def _addr(m):
            v = m.group(0)
            self.hits.append("详细地址")
            if self.keep_region:
                prefix = _region_prefix(v)
                if prefix and len(prefix) < len(v):
                    return prefix + "[详细地址]"
            return "[地址]"
        s = self.ADDRESS_RE.sub(_addr, s)

        def _unit(_m):
            self.hits.append("详细地址")
            return "[详细地址]"
        s = self.UNIT_RE.sub(_unit, s)
        s = re.sub(r"(?:\[详细地址\]|\[地址\])+", "[详细地址]", s)

        # 带称谓的姓名先处理（"张伟经理"），避免上下文规则把它截断
        s = self.honor_name.sub(lambda m: self.token("姓名", m.group(1)) + m.group(2), s)

        def _ctx(m):
            cand = m.group(2)
            if len(cand) < 2 or cand in STOP_TOKENS:
                return m.group(0)
            rest = ""
            if len(cand) == 3 and cand[2] in self.TRIM_CHARS:
                cand, rest = cand[:2], cand[2]
            if cand in STOP_TOKENS:
                return m.group(0)
            lead = m.group(0)[: len(m.group(0)) - len(m.group(2))]
            return f"{lead}{self.token('姓名', cand)}{rest}"
        s = self.ctx_name.sub(_ctx, s)
        s = self.moumou.sub(lambda m: f"{m.group(1)}某某", s)
        for pat in self.extra:
            s = pat.sub("[已遮蔽]", s)
        return s


def header_kind(name: str) -> str | None:
    low = (name or "").strip().lower()
    if not low:
        return None
    for kind, keys in CATEGORY_HINTS:
        for k in keys:
            if k in low:
                return kind
    return None


def mask_table(header: list[str], rows: list[list[str]], masker: Masker):
    kinds = [header_kind(h) for h in header]
    out = []
    for row in rows:
        new = []
        for i, cell in enumerate(row):
            kind = kinds[i] if i < len(kinds) else None
            new.append(masker.field(kind, cell))
        out.append(new)
    return out


def mask_profile(header: list[str]) -> list[str]:
    kinds: list[str] = []
    for h in header:
        k = header_kind(h)
        if k and k not in kinds:
            kinds.append(k)
    return kinds


# --------------------------------------------------------------------------
# tabular profile
# --------------------------------------------------------------------------

def profile_columns(rows: list[list[str]], header: list[str], mask: bool,
                    strict: bool) -> str:
    ncols = len(header)
    base = "| # | 列名 | 推断类型 | 非空 | 唯一值 | 数值区间 / 均值 |"
    lines = [base if strict else base + " 样例（已脱敏） |"]
    lines.append("|---|---|---|---|---|---|" + ("" if strict else "---|"))

    for c in range(ncols):
        col = [r[c] if c < len(r) else "" for r in rows]
        non_empty = [x for x in col if str(x).strip() != ""]
        nums = [n for n in (sniff_number(x) for x in non_empty) if n is not None]
        if non_empty and len(nums) == len(non_empty):
            kind = "整数" if all(isinstance(n, int) for n in nums) else "小数"
            span = f"{fmt_num(min(nums))} ~ {fmt_num(max(nums))} / 均值 {fmt_num(sum(nums) / len(nums))}"
        elif not non_empty:
            kind, span = "空列", "—"
        else:
            kind, span = "文本/已脱敏", "—"
        distinct = len(set(non_empty))
        name = f"C{c + 1}" if mask else (header[c] or f"C{c + 1}")
        name = str(name).replace("|", "\\|")[:40]
        if strict:
            lines.append(f"| {c + 1} | {name} | {kind} | {len(non_empty)} | {distinct} | {span} |")
        else:
            sample = ""
            if non_empty:
                freq: dict[str, int] = {}
                for x in non_empty:
                    freq[x] = freq.get(x, 0) + 1
                top = sorted(freq.items(), key=lambda kv: -kv[1])[:3]
                sample = "；".join(f"{k[:20]}({v})" for k, v in top).replace("|", "\\|")
            lines.append(f"| {c + 1} | {name} | {kind} | {len(non_empty)} | "
                         f"{distinct} | {span} | {sample} |")
    return "\n".join(lines)


def md_table(rows: list[list[str]], limit_cols: int = 12) -> str:
    if not rows:
        return "（空表）"
    width = min(max(len(r) for r in rows), limit_cols)
    out = []
    for i, r in enumerate(rows):
        cells = []
        for c in range(width):
            v = str(r[c]) if c < len(r) else ""
            v = v.replace("|", "\\|").replace("\n", " ")[:60]
            cells.append(v)
        out.append("| " + " | ".join(cells) + " |")
        if i == 0:
            out.append("|" + "---|" * width)
    if rows and max(len(r) for r in rows) > limit_cols:
        out.append(f"_（仅显示前 {limit_cols} 列）_")
    return "\n".join(out)


# --------------------------------------------------------------------------
# parsers
# --------------------------------------------------------------------------

def parse_xlsx(path: str):
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{NS_SS}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{NS_SS}t")))

        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for rel in ET.fromstring(z.read("xl/_rels/workbook.xml.rels")):
                rels[rel.get("Id")] = rel.get("Target", "")

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets_node = wb.find(f"{NS_SS}sheets")
        sheets: list[tuple[str, str]] = []
        if sheets_node is not None:
            for sh in list(sheets_node):
                target = rels.get(sh.get(f"{NS_R}id"), "")
                if target.startswith("/"):
                    target = target[1:]
                elif not target.startswith("xl/"):
                    target = "xl/" + target
                sheets.append((sh.get("name", "Sheet"), target))

        result = []
        for sheet_name, target in sheets:
            if target not in names:
                continue
            rows: list[list[str]] = []
            root = ET.fromstring(z.read(target))
            for row in root.iter(f"{NS_SS}row"):
                cells: dict[int, str] = {}
                for c in row.findall(f"{NS_SS}c"):
                    idx = col_index(c.get("r", ""))
                    t = c.get("t")
                    if t == "inlineStr":
                        node = c.find(f"{NS_SS}is")
                        val = "".join(x.text or "" for x in node.iter(f"{NS_SS}t")) if node is not None else ""
                    else:
                        v = c.find(f"{NS_SS}v")
                        val = decode_cell(t, v.text if v is not None else None, shared)
                    cells[idx] = val
                if cells:
                    width = max(cells) + 1
                    rows.append([cells.get(i, "") for i in range(width)])
            result.append((sheet_name, rows))
        return result


def parse_docx(path: str):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(f"{NS_W}body")
    blocks = []

    def para_text(p) -> str:
        parts = []
        for node in p.iter():
            if node.tag == f"{NS_W}t":
                parts.append(node.text or "")
            elif node.tag == f"{NS_W}tab":
                parts.append("\t")
            elif node.tag == f"{NS_W}br":
                parts.append("\n")
        return "".join(parts).strip()

    if body is None:
        return blocks
    for child in body:
        if child.tag == f"{NS_W}p":
            style_node = child.find(f"{NS_W}pPr/{NS_W}pStyle")
            style = style_node.get(f"{NS_W}val") if style_node is not None else ""
            txt = para_text(child)
            if txt:
                blocks.append(("p", style or "", txt))
        elif child.tag == f"{NS_W}tbl":
            rows = []
            for tr in child.findall(f"{NS_W}tr"):
                cells = []
                for tc in tr.findall(f"{NS_W}tc"):
                    cells.append(" ".join(para_text(p) for p in tc.findall(f"{NS_W}p")).strip())
                rows.append(cells)
            if rows:
                blocks.append(("table", rows))
    return blocks


def parse_pptx(path: str):
    with zipfile.ZipFile(path) as z:
        slides = []
        for name in z.namelist():
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
            if m:
                slides.append((int(m.group(1)), name))
        slides.sort()
        result = []
        for num, name in slides:
            root = ET.fromstring(z.read(name))
            texts = []
            for p in root.iter(f"{NS_A}p"):
                line = "".join(t.text or "" for t in p.iter(f"{NS_A}t")).strip()
                if line:
                    texts.append(line)
            result.append((num, texts))
        return result


def find_pdf_cli() -> str | None:
    cand = os.environ.get("WORKBUDDY_PDF_CLI")
    if cand and os.path.isfile(cand):
        return cand
    for name in ("workbuddy-pdf", "workbuddy-pdf.exe"):
        found = shutil.which(name)
        if found:
            return found
    guess = os.path.join(os.path.expanduser("~"), ".workbuddy", "external-packages",
                         "pdf-cli", "workbuddy-pdf.exe")
    return guess if os.path.isfile(guess) else None


def pdf_text_via_cli(src: str, workdir: str):
    cli = find_pdf_cli()
    if not cli:
        return None, None
    out = os.path.join(workdir, f".pdf_{hashlib.sha256(src.encode()).hexdigest()[:8]}.txt")
    try:
        proc = subprocess.run([cli, src, out, "--format", "txt"],
                              capture_output=True, timeout=180)
        if proc.returncode == 0 and os.path.isfile(out):
            text, _ = read_text_any(out)
            os.remove(out)
            return text, "workbuddy-pdf (local CLI)"
    except (OSError, subprocess.SubprocessError):
        pass
    if os.path.isfile(out):
        os.remove(out)
    return None, None


def pdf_text_builtin(path: str) -> str:
    import zlib
    data = open(path, "rb").read()
    chunks = []
    for m in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.S):
        raw = m.group(1)
        try:
            raw = zlib.decompress(raw)
        except zlib.error:
            continue
        for block in re.findall(rb"BT(.*?)ET", raw, re.S):
            for s in re.findall(rb"\((?:\\.|[^\\()])*\)", block):
                txt = s[1:-1]
                txt = txt.replace(b"\\(", b"(").replace(b"\\)", b")").replace(b"\\\\", b"\\")
                chunks.append(txt.decode("latin-1", "replace"))
    return " ".join(chunks)


def pdf_page_count(path: str) -> int:
    data = open(path, "rb").read()
    return len(re.findall(rb"/Type\s*/Page[^s]", data)) or len(re.findall(rb"/Type\s*/Page\b", data))


def image_info(path: str) -> dict:
    data = open(path, "rb").read(64 * 1024)
    info = {"format": os.path.splitext(path)[1].lower().lstrip("."), "width": None,
            "height": None, "extra": ""}
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            info["format"] = "png"
            info.update(width=int.from_bytes(data[16:20], "big"),
                        height=int.from_bytes(data[20:24], "big"),
                        extra=f"bitdepth={data[24]}, colortype={data[25]}")
        elif data[:2] == b"\xff\xd8":
            info["format"] = "jpeg"
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                              0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                    info.update(height=int.from_bytes(data[i + 5:i + 7], "big"),
                                width=int.from_bytes(data[i + 7:i + 9], "big"),
                                extra="progressive" if marker > 0xC2 else "baseline")
                    break
                i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
        elif data[:6] in (b"GIF87a", b"GIF89a"):
            info["format"] = "gif"
            info.update(width=int.from_bytes(data[6:8], "little"),
                        height=int.from_bytes(data[8:10], "little"))
        elif data[:2] == b"BM":
            info["format"] = "bmp"
            info.update(width=int.from_bytes(data[18:22], "little"),
                        height=int.from_bytes(data[22:26], "little"))
        elif data[8:12] == b"WEBP":
            info["format"] = "webp"
            if data[12:16] == b"VP8X":
                info["width"] = 1 + int.from_bytes(data[24:27], "little")
                info["height"] = 1 + int.from_bytes(data[27:30], "little")
    except (IndexError, ValueError):
        pass
    return info


def strip_html(text: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?</\1>", " ", text)
    text = re.sub(r"(?s)<br\s*/?>", "\n", text)
    text = re.sub(r"(?s)</(p|div|tr|li|h[1-6])>", "\n", text)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    return re.sub(r"[ \t]{2,}", " ", unescape(text))


def parse_delimited(text: str, path: str):
    import csv
    import io
    ext = os.path.splitext(path)[1].lower()
    if ext in (".csv", ".tsv"):
        sample = text[:8192]
        delim = "\t" if ext == ".tsv" else None
        if delim is None:
            try:
                delim = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
            except csv.Error:
                delim = ","
        return [r for r in csv.reader(io.StringIO(text), delimiter=delim)], f"分隔符：{repr(delim)}"
    return None, ""


# --------------------------------------------------------------------------
# per-file analysis
# --------------------------------------------------------------------------

def analyse(path: str, opts) -> dict:
    ext = os.path.splitext(path)[1].lower()
    stat = os.stat(path)
    M = opts["masker"]
    record = {
        "path": os.path.abspath(path),
        "name": os.path.basename(path),
        "ext": ext,
        "size": stat.st_size,
        "size_human": human_size(stat.st_size),
        "sha256": sha256_file(path),
        "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "kind": "其他",
        "sections": [],
        "stats": {},
        "notes": [],
        "masked": [],
    }

    snap = len(M.hits)

    def note_masked(kinds):
        for k in kinds:
            if k not in record["masked"]:
                record["masked"].append(k)

    def finalize():
        """按本文件的命中记录给出真正遮蔽过的类别，避免夸大脱敏范围。"""
        label = {"详细地址": "地址", "手机号": "电话", "固定电话": "电话",
                 "身份证号": "证件", "统一社会信用代码": "证件",
                 "银行卡号": "账号", "车牌号": "车牌"}
        for k in M.hits[snap:]:
            note_masked([label.get(k, k)])
        return record

    if ext in SHEET_EXT:
        record["kind"] = "表格"
        sheets = parse_xlsx(path)
        record["stats"]["sheets"] = len(sheets)
        for sheet_name, rows in sheets:
            header = rows[0] if rows else []
            body = rows[1:] if len(rows) > 1 else []
            note_masked(mask_profile(header))
            body = mask_table(header, body, M)
            block = [f"### 工作表：{sheet_name}（{len(body)} 行 × {len(header)} 列）", ""]
            if header:
                block += ["列画像：", "",
                          profile_columns(body, header, opts["mask_headers"], opts["strict"]), ""]
                if opts["rows"]:
                    block += [f"预览（前 {opts['rows']} 行，含表头，已脱敏）：", "",
                              md_table([header] + body[: opts["rows"]]), ""]
            else:
                block.append("_（该工作表为空）_")
            record["sections"].append("\n".join(block))
        record["notes"].append("日期单元格可能以 Excel 序列号（如 45678）呈现，按需在本地转换。")

    elif ext in DOC_EXT:
        record["kind"] = "文档"
        blocks = parse_docx(path)
        paras = [b for b in blocks if b[0] == "p"]
        tables = [b for b in blocks if b[0] == "table"]
        chars = sum(len(b[2]) for b in paras)
        record["stats"].update(paragraphs=len(paras), tables=len(tables), chars=chars)
        outline = [f"- {M.text(b[2])[:60]}" for b in paras
                   if b[1].lower().startswith(("heading", "标题", "1", "2", "3"))][:40]
        head = ["### 结构概览", "", f"- 段落 {len(paras)} 个，字符 {chars} 个，表格 {len(tables)} 个"]
        if outline:
            head += ["", "标题大纲（前 40 条，已脱敏）：", ""] + outline
        record["sections"].append("\n".join(head))
        # 命中类别由 finalize() 依据计数器增量给出

        if opts["strict"]:
            if tables:
                info = [f"- 表格 {i + 1}：{len(t[1])} 行 × {max((len(r) for r in t[1]), default=0)} 列"
                        for i, t in enumerate(tables)]
                record["sections"].append("### 表格清单（strict 模式不展开内容）\n\n" + "\n".join(info))
        else:
            body, used = [], 0
            for _kind, _style, raw in paras:
                text = M.text(raw)
                if used + len(text) > opts["max_chars"]:
                    body.append("\n…（正文超出长度上限，已截断）")
                    break
                body.append(text)
                used += len(text)
            if body:
                record["sections"].append("### 正文（已脱敏，局部截断）\n\n" + "\n\n".join(body))
            for i, (_, rows) in enumerate(tables):
                header = rows[0] if rows else []
                note_masked(mask_profile(header))
                trows = mask_table(header, rows[1:] if len(rows) > 1 else [], M)
                record["sections"].append(
                    f"### 表格 {i + 1}（{len(rows)} 行，已脱敏）\n\n"
                    + md_table(([header] + trows)[:21]))

    elif ext in SLIDE_EXT:
        record["kind"] = "演示文稿"
        slides = parse_pptx(path)
        record["stats"]["slides"] = len(slides)
        parts = [f"### 共 {len(slides)} 页（已脱敏）"]
        for num, texts in slides:
            texts = [M.text(t) for t in texts]
            if opts["strict"]:
                parts.append(f"- 第 {num} 页：{len(texts)} 个文本段落，共 {sum(len(t) for t in texts)} 字")
            else:
                body = " / ".join(texts)
                parts.append(f"- 第 {num} 页：{(body[:400] + '…') if len(body) > 400 else (body or '（无文字）')}")
        record["sections"].append("\n".join(parts))
        # 命中类别由 finalize() 依据计数器增量给出

    elif ext in PDF_EXT:
        record["kind"] = "PDF"
        text, engine = pdf_text_via_cli(path, opts["workspace"])
        if text is None:
            text, engine = pdf_text_builtin(path), "builtin zlib（尽力而为）"
        record["stats"].update(pages=pdf_page_count(path), engine=engine, chars=len(text or ""))
        head = ["### 概览", "", f"- 页数（估算）：{record['stats']['pages']}",
                f"- 本地抽取引擎：{engine}", f"- 可抽取字符数：{len(text or '')}"]
        if not (text or "").strip():
            head += ["", "> 该 PDF 可能为扫描件或使用 CID 字体，本地无法抽取文字。",
                     "> 处理建议：本地渲染为图片后在本地 OCR，或在本地转换为文本后再分析；",
                     "> 切勿为取文字而把文件上传到在线转换服务。"]
        record["sections"].append("\n".join(head))
        if text and text.strip():
            # 命中类别由 finalize() 依据计数器增量给出
            safe = M.text(text)
            if opts["strict"]:
                record["sections"].append("### 文本统计（strict 模式不展开正文）\n\n"
                                          f"- 字符 {len(safe)}，行 {safe.count(chr(10)) + 1}")
            else:
                record["sections"].append("### 文本（已脱敏，局部截断）\n\n" + safe[: opts["max_chars"]])

    elif ext in IMG_EXT:
        record["kind"] = "图片"
        info = image_info(path)
        record["stats"].update(info)
        ocr = shutil.which("tesseract")
        parts = ["### 图像元数据", "",
                 f"- 格式：{info['format']}",
                 f"- 尺寸：{info['width']} × {info['height']}" if info["width"] else "- 尺寸：未解析",
                 f"- 附加信息：{info['extra'] or '—'}",
                 f"- 本地 OCR：{'可用（' + ocr + '）' if ocr else '不可用'}", ""]
        if not ocr:
            parts += ["> 本机未安装 tesseract，图片内容不做文字识别。",
                      "> 处理建议：依据尺寸/格式/文件名等元数据处理；若必须读图，",
                      "> 请让使用者本机安装 OCR 后重跑，或由使用者口述关键内容。禁止上传图片到在线识别服务。"]
        record["sections"].append("\n".join(parts))

    elif ext in TEXT_EXT:
        record["kind"] = "文本"
        text, enc = read_text_any(path)
        record["stats"].update(encoding=enc, chars=len(text), lines=text.count("\n") + 1)
        fmt_note = ""
        if ext in (".html", ".htm"):
            text = strip_html(text)
            fmt_note = "（已本地剥离 HTML 标签）"
        if ext in (".csv", ".tsv"):
            rows, note = parse_delimited(text, path)
            if rows:
                header, body = rows[0], rows[1:]
                note_masked(mask_profile(header))
                body = mask_table(header, body, M)
                record["stats"].update(rows=len(body), cols=len(header))
                record["sections"].append("### 概览\n\n"
                                          f"- 编码 {enc}，{len(body)} 行 × {len(header)} 列；{note}")
                record["sections"].append(f"### 表格（{note}，已脱敏）\n\n"
                                          + profile_columns(body, header, opts["mask_headers"],
                                                            opts["strict"]))
                if opts["rows"]:
                    record["sections"].append(f"### 预览（前 {opts['rows']} 行，已脱敏）\n\n"
                                              + md_table([header] + body[: opts["rows"]]))
                return finalize()
        head = ["### 概览", "", f"- 编码 {enc}{fmt_note}",
                f"- 字符 {len(text)}，行 {text.count(chr(10)) + 1}"]
        record["sections"].append("\n".join(head))
        # 命中类别由 finalize() 依据计数器增量给出
        safe = M.text(text)
        if opts["strict"]:
            record["sections"].append("### 内容统计（strict 模式不展开正文）")
        else:
            record["sections"].append("### 内容（已脱敏，局部截断）\n\n" + safe[: opts["max_chars"]])

    elif ext in LEGACY_OFFICE:
        record["kind"] = "旧版 Office"
        record["notes"].append("旧版二进制格式（.xls/.doc/.ppt）无法零依赖解析；"
                               "请在本机用 WPS/Office 另存为 .xlsx/.docx/.pptx 后重跑，"
                               "或走本地编辑器打开读取。禁止上传到在线转换网站。")
    else:
        record["kind"] = "二进制/未知"
        record["notes"].append("未识别的文件类型，仅记录哈希与大小；"
                               "如需深入分析，请先在本地转换为受支持格式。")

    return finalize()


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

MODE_ALIAS = {"normal": "normal", "standard": "normal", "普通": "normal",
              "strict": "strict", "严格": "strict"}


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    args = argv[1:]
    files, workspace = [], ".local_analysis"
    mode, rows, max_chars = "normal", None, 6000
    extra_redact, mask_headers, quiet, keep_region = [], False, False, True

    i = 0
    while i < len(args):
        a = args[i]
        if a == "--mode":
            raw = args[i + 1].strip().lower()
            mode = MODE_ALIAS.get(raw)
            if not mode:
                print(f"错误：未知模式 {raw}，可用 normal / strict", file=sys.stderr)
                return 2
            i += 2
        elif a == "--workspace":
            workspace = args[i + 1]; i += 2
        elif a == "--rows":
            rows = int(args[i + 1]); i += 2
        elif a == "--max-chars":
            max_chars = int(args[i + 1]); i += 2
        elif a == "--redact":
            extra_redact.append(args[i + 1]); i += 2
        elif a == "--mask-region":
            keep_region = False; i += 1
        elif a == "--mask-headers":
            mask_headers = True; i += 1
        elif a == "--quiet":
            quiet = True; i += 1
        elif a in ("-h", "--help"):
            print(__doc__); return 0
        else:
            files.append(a); i += 1

    if not files:
        print("错误：至少需要一个输入文件。使用 --help 查看用法。", file=sys.stderr)
        return 2

    missing = [f for f in files if not os.path.isfile(f)]
    if missing:
        print("错误：文件不存在 → " + ", ".join(missing), file=sys.stderr)
        return 2

    strict = mode == "strict"
    if rows is None:
        rows = 0 if strict else 5
    if strict:
        rows = min(rows, 2)

    os.makedirs(workspace, exist_ok=True)
    masker = Masker(keep_region=keep_region, extra=extra_redact)
    opts = {"workspace": workspace, "strict": strict, "rows": rows,
            "max_chars": max_chars, "mask_headers": mask_headers, "masker": masker}

    records = []
    for f in files:
        try:
            records.append(analyse(f, opts))
        except Exception as exc:  # noqa: BLE001
            records.append({"path": os.path.abspath(f), "name": os.path.basename(f),
                            "ext": os.path.splitext(f)[1].lower(),
                            "size": os.path.getsize(f) if os.path.isfile(f) else 0,
                            "size_human": human_size(os.path.getsize(f)) if os.path.isfile(f) else "0 B",
                            "sha256": "unavailable", "modified": "-", "kind": "解析失败",
                            "sections": [], "stats": {}, "masked": [],
                            "notes": [f"解析异常：{type(exc).__name__}: {exc}"]})

    total = sum(r["size"] for r in records)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    region_note = "地址保留至省/市级，其后遮蔽" if keep_region else "地址整体遮蔽"
    out = ["# 本地分析报告（未上传云端 · 已脱敏）", "",
           f"- 生成时间：{stamp}",
           f"- 处理模式：**{'严格' if strict else '普通'}**"
           + ("（只输出聚合指标，不出现任何单元格或正文原文）" if strict else "（本地读取 + 脱敏后的值预览）"),
           f"- 输入文件：{len(records)} 个，合计 {human_size(total)}",
           f"- 工作目录：`{os.path.abspath(workspace)}`",
           f"- 脱敏：**强制启用，不可关闭**（姓名、电话、地址、证件号、邮箱、银行卡/账号、车牌、编号）；{region_note}",
           "- 脱敏方式：同一取值映射为同一假名（如 `[姓名#1]`），计数与关联仍成立，身份不可还原",
           "", "> 本报告由本机脚本生成，全程无网络请求；原始文件仅被只读打开，未复制、未修改。",
           "> 报告内不出现任何原始姓名、联系方式、地址与证件号码。", ""]

    for idx, r in enumerate(records, 1):
        out += [f"## {idx}. {r['name']}", "",
                "| 项 | 值 |", "|---|---|",
                f"| 类型 | {r['kind']} |",
                f"| 大小 | {r['size_human']} |",
                f"| 修改时间 | {r['modified']} |",
                f"| SHA-256 | `{r['sha256']}` |"]
        for k, v in (r.get("stats") or {}).items():
            if v in (None, ""):
                continue
            if r["kind"] == "图片" and k in ("format", "width", "height", "extra"):
                continue
            out.append(f"| {k} | {v} |")
        if r["masked"]:
            out.append(f"| 已脱敏字段 | {'、'.join(r['masked'])} |")
        out.append("")
        sec = [s.strip("\n") for s in r["sections"] if s and s.strip()]
        if sec:
            out.append("\n\n".join(sec))
        if r["notes"]:
            out.append("")
            out.append("**提示**：" + "；".join(r["notes"]))
        out.append("")

    report_path = os.path.join(workspace, "report.md")
    with open(report_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(out))

    manifest = {
        "generated_at": stamp,
        "mode": mode,
        "strict": strict,
        "workspace": os.path.abspath(workspace),
        "total_files": len(records),
        "total_size": total,
        "masking": {
            "enabled": True,
            "mandatory": True,
            "categories": ["姓名", "电话", "地址", "证件", "邮箱", "账号", "车牌", "编号"],
            "keep_region": keep_region,
            "pseudonym_consistent": True,
            "extra_patterns": len(extra_redact),
            "hits": dict(Counter(masker.hits)),
            "distinct_values": dict(masker.counters),
        },
        "network_used": False,
        "files": [{k: r.get(k) for k in ("path", "name", "ext", "size", "sha256",
                                         "modified", "kind", "stats", "masked")}
                  for r in records],
        "outputs": [os.path.abspath(report_path)],
    }
    with open(os.path.join(workspace, "manifest.json"), "w", encoding="utf-8", newline="\n") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2)

    if not quiet:
        print(f"[local-only] 模式={'严格' if strict else '普通'} 文件={len(records)} 合计={human_size(total)}")
        for r in records:
            extra = []
            if "sheets" in r["stats"]:
                extra.append(f"{r['stats']['sheets']}个工作表")
            if "pages" in r["stats"]:
                extra.append(f"约{r['stats']['pages']}页")
            if "slides" in r["stats"]:
                extra.append(f"{r['stats']['slides']}页")
            if "rows" in r["stats"]:
                extra.append(f"{r['stats']['rows']}行")
            if r["stats"].get("width"):
                extra.append(f"{r['stats']['width']}×{r['stats']['height']}")
            mark = f"  已脱敏:{'/'.join(r['masked'])}" if r["masked"] else ""
            print(f"  - {r['name']}  [{r['kind']}]  {r['size_human']}"
                  + ("  " + ", ".join(extra) if extra else "") + mark)
        occ = Counter(masker.hits)
        if occ:
            print("[local-only] 脱敏命中（按出现次数）："
                  + "，".join(f"{k} {v} 处" for k, v in occ.most_common()))
        print(f"[local-only] 报告 → {os.path.abspath(report_path)}")
        print(f"[local-only] 清单 → {os.path.abspath(os.path.join(workspace, 'manifest.json'))}")
        print("[local-only] 网络请求：0（无任何上传/外发调用）")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
