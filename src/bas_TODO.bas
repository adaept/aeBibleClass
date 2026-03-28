Attribute VB_Name = "bas_TODO"
Option Explicit

' Embedded Extension Hooks (Implicit but Intentional)
'=====================================================
' Language variants > alternate BookName lexemes
' Pericope titles / version tags > append after Reference

'   Next more advanced refinement would be:
'     - Precompute the verse maps once
'     - Cache them in a module-level structure
'     - Make lookup entirely allocation-free

'Stage 16 - Verse Expansion

'multi-book lists
'verse lists inside chapters
'chapter shorthand (John 3,4,5)
'OSIS Export
'packed verse ID expansion

'5. Optional Defensive Test Stage 9 (Future)
'
'Not necessary now, but eventually it is useful to protect against a classic bug:
'
'1 -2 Samuel
'
'A naive range detector may interpret this as a range.
'Later you may add:
'PASS: book prefix dash not treated As Word.Range
'But this can also be handled naturally once Stage 2 tokenization is used during composition, so it's not urgent.
'================================================================================================================

'Stage-8: Reference Expansion Engine
'
'It converts canonical references into explicit verse ranges so downstream systems (search, highlighting, indexing, cross-references) can work deterministically.
'
'Stage-8: Verse Expansion Engine
'Purpose
'
'Convert any canonical reference into a fully enumerated verse list.
'
'Example inputs from your Stage-7 output:
'
'Romans 8
'Romans 8:28
'Romans 8:28-30
'Romans 8:28,30
'Genesis 1 - 3
'
'Expanded internal representation:
'
'Romans 8:1
'Romans 8:2
'Romans 8:3
'...
'Romans 8:39
'
'or
'
'Romans 8:28
'Romans 8:29
'Romans 8:30
'Why Bible Software Uses This
'
'This enables:
'
'1 - Fast search hit marking
'Highlight verses directly.
'
'2 - Cross-reference linking
'Jump to exact verse.
'
'3 - Verse range math
'Union / intersection of references.
'
'4 - Consistent indexing
'Every verse has a unique numeric ID.
'
'Internal Representation Used by Many Systems
'
'Professional engines convert to a VerseID:
'
'VerseID = BookID * 1,000,000 + Chapter * 1,000 + Verse
'
'Example:
'
'Genesis 1:1  -> 1001001
'Romans 8:28  -> 45008028 (example depending on scheme)
'
'This allows:
'
'range comparisons
'sorting
'fast lookup
'Stage-8 Architecture
'User Input
'     -
'Stage 1  Normalize
'Stage 2  Lexical Scan
'Stage 3  Resolve Alias
'Stage 4  Interpret Structure
'Stage 5  Validate
'Stage 6  Canonical Format
'Stage 7  End-to-End Parse
'
'     Print
'Stage 8  Expand to Verse Set
'
'Output Example:
'
'Romans 8
'
'becomes:
'
'bookID = 45
'chapter = 8
'verses = [1..39]
'VBA Implementation(Stage - 8)
'Verse Expansion Function
'Public Function ExpandReference(bookID As Long, chapter As Long, verseSpec As String) As Collection
'
'    Dim verses As New Collection
'    Dim parts() As String
'    Dim i As Long
'    Dim vStart As Long
'    Dim vEnd As Long
'    Dim v As Long
'
'    If verseSpec = "" Then
'
'        ' Entire chapter
'        vEnd = GetMaxVerse(bookID, chapter)
'
'        For v = 1 To vEnd
'            verses.Add v
'        Next v
'
'        Set ExpandReference = verses
'        Exit Function
'
'    End If
'
'    parts = Split(verseSpec, ",")
'
'    For i = LBound(parts) To UBound(parts)
'
'        If InStr(parts(i), "-") > 0 Then
'
'            vStart = CLng(Split(parts(i), "-")(0))
'            vEnd = CLng(Split(parts(i), "-")(1))
'
'            For v = vStart To vEnd
'                verses.Add v
'            Next v
'
'        Else
'
'            verses.Add CLng(parts(i))
'
'        End If
'
'    Next i
'
'    Set ExpandReference = verses
'
'End Function
'Example
'
'Input
'
'Romans 8:28-30,32
'
'Result:
'
'28
'29
'30
'32
'Even More Powerful (Used by Logos / Accordance)
'
'They expand to VerseID ranges:
'
'StartVerseID
'EndVerseID
'
'Example:
'
'Romans 8:28-30
'
'becomes
'
'45008028
'45008030
'
'This makes range comparisons extremely fast.
'

'Phase 2 - Performance Layer
'
'Now you optimize.
'
'Stage 16 - Zero-Allocation Packed Engine
'
'Convert packed verse map into:
'
'bitset array (66 books)
'1189 Chapter offsets
'constant-time lookup
'
'This gives:
'
'O(1) verse lookup
'O(1) intersection
'O(1) union
'instant highlighting
'
'This is how Logos / Accordance class engines work.
