"""
fix_abc.py — Apply fixes A, B, C to aeBibleClass.cls

A: Remove dead totalParaCount variable from CountEmptyParasWithNoThemeColor
B: Change all As Integer -> As Long
C: Add m_lastFuncError = True before every Resume PROC_EXIT
"""

import re

path = r'C:\adaept\aeBibleClass\src\aeBibleClass.cls'

with open(path, 'r', encoding='utf-8-sig') as f:
    src = f.read()

original_len = len(src)

# ── Fix B ──────────────────────────────────────────────────────────────────
# Change all As Integer -> As Long (return types, Dim vars, parameters)
src = src.replace(' As Integer', ' As Long')

b_count = src.count(' As Long')  # for verification

# ── Fix A ──────────────────────────────────────────────────────────────────
# Remove dead totalParaCount Dim line
src = re.sub(r'    Dim totalParaCount As Long\r?\n', '', src)
# Remove dead totalParaCount assignment (line + blank line that follows)
src = re.sub(r'    totalParaCount = ActiveDocument\.Paragraphs\.Count\r?\n    \r?\n', '\r\n', src)

# ── Fix C — declaration (insert after m_ReportBuf line) ───────────────────
OLD_DECL = "Private m_ReportBuf As String          ' Accumulates TestReport lines; flushed once at end of run"
NEW_DECL = (OLD_DECL + "\r\n"
            "Private m_lastFuncError As Boolean     "
            "' Set True in any Count function PROC_ERR; checked in GetPassFail")
src = src.replace(OLD_DECL, NEW_DECL, 1)

# ── Fix C — set flag before every Resume PROC_EXIT ───────────────────────
# Pattern: four-space indent Resume PROC_EXIT
src = src.replace('    Resume PROC_EXIT',
                  '    m_lastFuncError = True\r\n    Resume PROC_EXIT')

resume_count = src.count('m_lastFuncError = True')

# ── Write result ──────────────────────────────────────────────────────────
with open(path, 'w', encoding='utf-8') as f:
    f.write(src)

print(f"Fix B: 'As Long' occurrences now = {b_count}")
print(f"Fix A: totalParaCount lines removed (check manually)")
print(f"Fix C: m_lastFuncError = True inserted {resume_count} times")
print(f"File length: {original_len} -> {len(src)} chars")
