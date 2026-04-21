"""
step4_hints.py — Infrastructure plan Step 4: first-hit hint arrays + RunTest print hook

Changes:
  1. Adds m_lastHint As String and m_HintArray(1 To MaxTests) As String declarations
  2. Resets both in InitializeGlobalResultArrayToMinusOne
  3. Adds m_HintArray(TestNum) = m_lastHint after every ResultArray(TestNum) = ... in GetPassFail
  4. Adds hint print line after every Debug.Print GetPassFailArray(num) line in RunTest

No Count functions are modified — m_lastHint is always "" until Step 5.
All hint lines will be silent on this run; RunTest output unchanged.
"""

import re

path = r'C:\adaept\aeBibleClass\src\aeBibleClass.cls'

with open(path, 'r', encoding='utf-8-sig') as f:
    src = f.read()

# ── 1. Declarations ────────────────────────────────────────────────────────
OLD_DECL = "Private m_lastFuncError As Boolean     ' Set True in any Count function PROC_ERR; checked in GetPassFail"
NEW_DECL = (
    "Private m_lastFuncError As Boolean     ' Set True in any Count function PROC_ERR; checked in GetPassFail\r\n"
    "Private m_lastHint As String           ' Set by Count functions on first violation; cleared in GetPassFail\r\n"
    "Private m_HintArray(1 To MaxTests) As String  ' Per-test first-hit hint; printed in RunTest on FAIL"
)
assert OLD_DECL in src, "Declaration anchor not found"
src = src.replace(OLD_DECL, NEW_DECL, 1)

# ── 2. Reset in InitializeGlobalResultArrayToMinusOne ─────────────────────
OLD_RESET = "    m_ReportBuf = \"\""
NEW_RESET = (
    "    m_ReportBuf = \"\"\r\n"
    "    m_lastHint = \"\"\r\n"
    "    Dim h As Long\r\n"
    "    For h = 1 To MaxTests : m_HintArray(h) = \"\" : Next h"
)
assert OLD_RESET in src, "Reset anchor not found"
src = src.replace(OLD_RESET, NEW_RESET, 1)

# ── 3. m_HintArray capture after every ResultArray assignment in GetPassFail
# Pattern: 8-space indent ResultArray(TestNum) = <anything>
# Add m_HintArray(TestNum) = m_lastHint on the next line
src = re.sub(
    r'(        ResultArray\(TestNum\) = [^\r\n]+)',
    lambda m: m.group(0) + '\r\n        m_HintArray(TestNum) = m_lastHint',
    src
)

# ── 4. Hint print after every Debug.Print GetPassFailArray(num) in RunTest ─
# Pattern: 8-space indent Debug.Print GetPassFailArray(num), ...
HINT_PRINT = ('        If GetPassFailArray(num) = "FAIL!!!!" '
              'And m_HintArray(num) <> "" Then '
              'Debug.Print , , , "  >> First hit: " & m_HintArray(num)')
src = re.sub(
    r'(        Debug\.Print GetPassFailArray\(num\)[^\r\n]+)',
    lambda m: m.group(0) + '\r\n' + HINT_PRINT,
    src
)

with open(path, 'w', encoding='utf-8') as f:
    f.write(src)

# ── Verify ─────────────────────────────────────────────────────────────────
hint_captures = src.count('m_HintArray(TestNum) = m_lastHint')
hint_prints   = src.count('>> First hit:')
print(f"m_HintArray(TestNum) = m_lastHint  inserted: {hint_captures}")
print(f"Hint print lines inserted:                   {hint_prints}")
print("Done.")
