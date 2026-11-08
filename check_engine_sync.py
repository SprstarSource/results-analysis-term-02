#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_engine_sync.py
====================
يتحقق من أن النسخ المضمّنة من «محرك التجميع» داخل الموديلات متطابقة حرفياً.

الخلفية: كل موديل مستقل يحمل نسخته الخاصة من المحرك (الإجراءات التي تبدأ بـ Src
وكل إجراء مساعد يشاركها فيه أكثر من موديل). التعليق في رأس كل موديل ينص على
"keep identical"، وهذا السكربت يجعل هذا الشرط قابلاً للفحص الآلي.

الاستعمال:
    python tools/check_engine_sync.py              # يفحص src/vba/*.bas
    python tools/check_engine_sync.py path/to/*.bas

رمز الخروج: 0 = متطابقة، 1 = اختلاف، 2 = خطأ استعمال.
"""
import glob
import re
import sys
from collections import defaultdict

DEFAULT_GLOB = "src/vba/*.bas"
# الموديلات التي تحمل نسخة مضمّنة من المحرك (تُستبعد البقية من المقارنة)
EMBEDDED = ("modAverage02", "modSplitTermResults", "modCompare", "modTransfer02", "modAnalysisBuilderData02")

# مساعدات مشتركة في الاسم لكنها خاصة بكل موديل عمداً (مجلد المخرجات) - لا تدخل في المقارنة
EXCLUDE = {"EnsureFolder", "OutputFolderName", "FolderExistsPath", "FileExistsPath"}

PROC_START = re.compile(r"^(?:Public |Private |Friend )?(?:Static )?(Sub|Function)\s+(\w+)", re.I)


def read(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    for enc in ("utf-8", "cp1256", "latin-1"):
        try:
            text = raw.decode(enc).replace("\r\n", "\n").replace("\r", "\n")
            return strip_local_extensions(text)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("encoding not supported: " + path)


# كتلة إضافية مقصودة في نسخة موديل واحد (مثل حفظ A1:A4 في modAnalysisBuilderData02)
# تُحاط بسطري تعليق ولا تدخل في المقارنة:
#     ' LOCAL-EXTENSION BEGIN ...
#     ' LOCAL-EXTENSION END
LOCAL_EXT = re.compile(r"^[ \t]*' LOCAL-EXTENSION BEGIN[^\n]*\n.*?^[ \t]*' LOCAL-EXTENSION END[^\n]*\n?", re.M | re.S)


def strip_local_extensions(text):
    return LOCAL_EXT.sub("", text)


def procedures(text):
    """name -> normalised text of the procedure (from its header to End Sub/Function)."""
    out, name, kind, buf = {}, None, None, []
    for line in text.split("\n"):
        if name is None:
            m = PROC_START.match(line.strip())
            if m:
                kind, name, buf = m.group(1).lower(), m.group(2), [line.rstrip()]
        else:
            buf.append(line.rstrip())
            if re.match(r"^\s*End\s+%s\b" % kind, line, re.I):
                out[name] = "\n".join(buf)
                name = None
    return out


def strip_comment(line):
    """يحذف تعليق VBA (') من نهاية السطر مع احترام النصوص بين علامتي اقتباس."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    return line


# أسماء ثوابت تختلف بين الموديلات لنفس القيمة (Source02)
ALIASES = {r"\bSRC_FOLDER\b": "SOURCE_FOLDER_NAME"}


def normalise(proc_text):
    """يطبّع الإجراء: بلا تعليقات ولا أسطر فارغة ولا فروق مسافات/إزاحة."""
    out = []
    for line in proc_text.split("\n"):
        code = re.sub(r"\s+", " ", strip_comment(line)).strip()
        for pat, rep in ALIASES.items():
            code = re.sub(pat, rep, code)
        if code:
            out.append(code)
    return "\n".join(out)


def module_name(text, path):
    m = re.search(r'^Attribute VB_Name\s*=\s*"(\w+)"', text, re.M)
    return m.group(1) if m else path


def main(argv):
    paths = argv or sorted(glob.glob(DEFAULT_GLOB))
    if not paths:
        print("لا توجد ملفات .bas للفحص:", DEFAULT_GLOB)
        return 2
    mods = {}
    for p in paths:
        t = read(p)
        n = module_name(t, p)
        if n in EMBEDDED:
            mods[n] = {k: normalise(v) for k, v in procedures(t).items() if k not in EXCLUDE}
    if len(mods) < 2:
        print("يلزم موديلان مضمَّنان على الأقل للمقارنة. وُجد:", sorted(mods))
        return 2

    holders = defaultdict(list)           # procedure name -> modules that define it
    for n, procs in mods.items():
        for p in procs:
            holders[p].append(n)
    engine = sorted(p for p, h in holders.items() if len(h) >= 2 and (p.startswith("Src") or len(h) == len(mods)))

    bad = 0
    for p in engine:
        variants = defaultdict(list)
        for n in holders[p]:
            variants[mods[n][p]].append(n)
        if len(variants) > 1:
            bad += 1
            print("✗ %-28s يختلف بين: %s" % (p, " | ".join(", ".join(v) for v in variants.values())))
    print("\nالموديلات المفحوصة :", ", ".join(sorted(mods)))
    print("إجراءات المحرك     : %d   المختلفة: %d" % (len(engine), bad))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
