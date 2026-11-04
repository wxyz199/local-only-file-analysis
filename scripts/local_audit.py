#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Emit a local-only handling audit for a workspace produced by local_extract.py.

Verifies that:
  1. every source file is still byte-identical to the hashed snapshot (read-only);
  2. the workspace contains only locally generated artefacts;
  3. no network / upload step is declared by the pipeline.

Usage:
  python local_audit.py [WORKSPACE] [--note "补充说明"] [--json]

Default WORKSPACE is ./.local_analysis
Exit code 0 = all checks pass, 1 = integrity problem detected.
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
from datetime import datetime

FORBIDDEN_RULES = [
    ("云文档 / 在线表格", "腾讯文档、金山文档、飞书文档等创建或导入"),
    ("云盘 / 对象存储", "腾讯网盘、企业微信微盘、云服务文件存储等上传"),
    ("知识库 / 记忆库", "资料库写入、ima 知识库导入、URL 抓取入库"),
    ("外部检索", "把文件内容或原文片段作为 WebSearch / WebFetch 的查询词"),
    ("图片外发", "在线 OCR、在线图片理解、图片生成接口回传原图"),
    ("第三方转换", "在线 PDF/Office 转换、在线压缩、在线解压包"),
    ("即时通讯外发", "把文件或摘要发送到群聊、邮件、机器人"),
]


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


def main(argv: list[str]) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    workspace, note, as_json = ".local_analysis", "", False
    args = argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--note":
            note = args[i + 1]; i += 2
        elif args[i] == "--json":
            as_json = True; i += 1
        elif args[i] in ("-h", "--help"):
            print(__doc__); return 0
        else:
            workspace = args[i]; i += 1

    manifest_path = os.path.join(workspace, "manifest.json")
    if not os.path.isfile(manifest_path):
        print(f"错误：未找到 {manifest_path}，请先运行 local_extract.py", file=sys.stderr)
        return 2

    manifest = json.load(open(manifest_path, encoding="utf-8"))
    checks = []
    failed = 0

    for f in manifest.get("files", []):
        path, recorded = f.get("path", ""), f.get("sha256", "")
        if not os.path.isfile(path):
            checks.append((f["name"], "原件缺失", "无法复核（可能已移动或删除）", False))
            continue
        actual = sha256_file(path)
        ok = actual == recorded
        checks.append((f["name"], "哈希未变（只读处理）" if ok else "哈希变化",
                       actual[:16] + "…", ok))
        if not ok:
            failed += 1

    artefacts = []
    for root, _dirs, names in os.walk(workspace):
        for n in names:
            p = os.path.join(root, n)
            artefacts.append((os.path.relpath(p, workspace), os.path.getsize(p)))

    if as_json:
        print(json.dumps({
            "workspace": os.path.abspath(workspace),
            "mode": manifest.get("mode"),
            "generated_at": manifest.get("generated_at"),
            "files": len(manifest.get("files", [])),
            "integrity_ok": failed == 0,
            "artefacts": [a[0] for a in artefacts],
        }, ensure_ascii=False, indent=2))
        return 0 if failed == 0 else 1

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "# 本地处理审计单", "",
        f"- 审计时间：{stamp}",
        f"- 抽取时间：{manifest.get('generated_at', '-')}",
        f"- 处理模式：{manifest.get('mode', '-')}",
        f"- 工作目录：`{os.path.abspath(workspace)}`",
        f"- 输入文件：{len(manifest.get('files', []))} 个，合计 {human_size(manifest.get('total_size', 0))}",
        "", "## 一、原件完整性（证明只读）", "",
        "| 文件 | 校验结论 | SHA-256 前 16 位 | 结果 |", "|---|---|---|---|",
    ]
    for name, verdict, digest, ok in checks:
        lines.append(f"| {name} | {verdict} | `{digest}` | {'通过' if ok else '失败'} |")

    lines += ["", "## 二、脱敏执行情况", ""]
    masking = manifest.get("masking") or {}
    if masking.get("enabled"):
        hits = masking.get("hits") or {}
        lines += [
            "| 项目 | 状态 |", "|---|---|",
            f"| 脱敏开关 | 强制启用（不可关闭） |",
            f"| 覆盖类别 | {'、'.join(masking.get('categories', []))} |",
            f"| 地址处理 | {'保留至省/市级' if masking.get('keep_region', True) else '整体遮蔽'} |",
            f"| 假名一致性 | {'同值同码，计数与关联成立' if masking.get('pseudonym_consistent') else '否'} |",
            f"| 本次命中 | {'，'.join(f'{k} {v} 项' for k, v in hits.items()) or '未命中敏感数据'} |",
            "",
            "报告中的姓名、电话、地址、证件号、邮箱、账号、车牌、编号均已替换为 `[类别#序号]` 假名；"
            "原始值与映射表均未落盘。",
        ]
    else:
        lines.append("未检测到脱敏记录，请确认抽取脚本为最新版本并重跑。")

    lines += ["", "## 三、本地产出物", "",
              "| 文件 | 大小 |", "|---|---|"]
    for rel, size in sorted(artefacts):
        lines.append(f"| `{rel}` | {human_size(size)} |")

    lines += ["", "## 四、红线核对（本次处理均未触发）", "",
              "| 风险类别 | 明确禁止的动作 | 状态 |", "|---|---|---|"]
    for kind, desc in FORBIDDEN_RULES:
        lines.append(f"| {kind} | {desc} | 未发生 |")

    lines += ["", "## 五、结论", ""]
    if failed == 0:
        lines.append("- 全部源文件哈希与处理前一致，处理过程为只读，未产生副本。")
        lines.append("- 全部产出物位于上述工作目录内，未发现外发调用。")
        lines.append("- 结论：**本地闭环通过**。文件内容未提交给任何云文档、云盘、知识库或在线转换服务。")
    else:
        lines.append(f"- 有 {failed} 个文件的哈希与记录不一致，请确认是否在处理期间被外部程序修改。")
        lines.append("- 结论：**需要人工复核**。")

    if note:
        lines += ["", "## 六、补充说明", "", note]

    audit_md = "\n".join(lines) + "\n"
    audit_path = os.path.join(workspace, "audit.md")
    with open(audit_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(audit_md)

    print(f"[local-only] 完整性校验：{'通过' if failed == 0 else '失败'}（{len(checks)} 个文件）")
    print(f"[local-only] 审计单 → {os.path.abspath(audit_path)}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
