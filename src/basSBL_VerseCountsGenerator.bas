Attribute VB_Name = "basSBL_VerseCountsGenerator"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

' Production Runtime Version
'   After generation, replace the dictionary entirely with:
'    - Private Function GetPackedVerseMap() As Variant
'       ...
'      End Function

' Output for the 66-book Protestant canon used by:
'    - King James Version
'    - New International Version
' Both contain:
'    - 66 books
'    - 1189 chapters
'    - 31,102 verses (KJV)
' (NIV verse structure is the same at the chapter level)

Private Function ExpectedChapterCounts() As Variant
    ExpectedChapterCounts = Array( _
        50, 40, 27, 36, 34, 24, 21, 4, 31, 24, _
        22, 25, 29, 36, 10, 13, 10, 42, 150, 31, _
        12, 8, 66, 52, 5, 48, 12, 14, 3, 9, _
        1, 4, 7, 3, 3, 3, 2, 14, 4, _
        28, 16, 24, 21, 28, 16, 16, 13, 6, 6, _
        4, 4, 5, 3, 6, 4, 3, 1, 13, 5, _
        5, 3, 5, 1, 1, 1, 22)
End Function

Private Function ToOneBasedLongArray(src As Variant, context As String) As Long()
    Dim count As Long
    count = UBound(src) - LBound(src) + 1
    
    Dim temp() As Long
    ReDim temp(1 To count)
    
    Dim i As Long
    Dim dest As Long
    dest = 1
    
    For i = LBound(src) To UBound(src)
        temp(dest) = CLng(src(i))
        dest = dest + 1
    Next i
    
    AssertOneBased temp, context
    
    ToOneBasedLongArray = temp
End Function

Public Function GetVerseCounts() As Object
' This table matches standard Protestant versification only.
' It Not does not match:
'    - LXX Psalms numbering
'    - Vulgate numbering
'    - Catholic additions
'    - Orthodox versification

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    ' -------- OLD TESTAMENT --------
    d.Add 1, ToOneBasedLongArray(Array(31, 25, 24, 26, 32, 22, 24, 22, 29, 32, 32, 20, 18, 24, 21, 16, 27, 33, 38, 18, 34, 24, 20, 67, 34, 35, 46, 22, 35, 43, 55, 32, 20, 31, 29, 43, 36, 30, 23, 23, 57, 38, 34, 34, 28, 34, 31, 22, 33, 26), "Genesis")
    d.Add 2, ToOneBasedLongArray(Array(22, 25, 22, 31, 23, 30, 25, 32, 35, 29, 10, 51, 22, 31, 27, 36, 16, 27, 25, 26, 36, 31, 33, 18, 40, 37, 21, 43, 46, 38, 18, 35, 23, 35, 35, 38, 29, 31, 43, 38), "Exodus")
    d.Add 3, ToOneBasedLongArray(Array(17, 16, 17, 35, 19, 30, 38, 36, 24, 20, 47, 8, 59, 57, 33, 34, 16, 30, 37, 27, 24, 33, 44, 23, 55, 46, 34), "Leviticus")
    d.Add 4, ToOneBasedLongArray(Array(54, 34, 51, 49, 31, 27, 89, 26, 23, 36, 35, 16, 33, 45, 41, 50, 13, 32, 22, 29, 35, 41, 30, 25, 18, 65, 23, 31, 40, 16, 54, 42, 56, 29, 34, 13), "Numbers")
    d.Add 5, ToOneBasedLongArray(Array(46, 37, 29, 49, 33, 25, 26, 20, 29, 22, 32, 32, 18, 29, 23, 22, 20, 22, 21, 20, 23, 30, 25, 22, 19, 19, 26, 68, 29, 20, 30, 52, 29, 12), "Deuteronomy")
    d.Add 6, ToOneBasedLongArray(Array(18, 24, 17, 24, 15, 27, 26, 35, 27, 43, 23, 24, 33, 15, 63, 10, 18, 28, 51, 9, 45, 34, 16, 33), "Joshua")
    d.Add 7, ToOneBasedLongArray(Array(36, 23, 31, 24, 31, 40, 25, 35, 57, 18, 40, 15, 25, 20, 20, 31, 13, 31, 30, 48, 25), "Judges")
    d.Add 8, ToOneBasedLongArray(Array(22, 23, 18, 22), "Ruth")
    d.Add 9, ToOneBasedLongArray(Array(28, 36, 21, 22, 12, 21, 17, 22, 27, 27, 15, 25, 23, 52, 35, 23, 58, 30, 24, 42, 15, 23, 29, 22, 44, 25, 12, 25, 11, 31, 13), "1 Samuel")
    d.Add 10, ToOneBasedLongArray(Array(27, 32, 39, 12, 25, 23, 29, 18, 13, 19, 27, 31, 39, 33, 37, 23, 29, 32, 44, 26, 22, 51, 39, 25), "2 Samuel")
    d.Add 11, ToOneBasedLongArray(Array(53, 46, 28, 34, 18, 38, 51, 66, 28, 29, 43, 33, 34, 31, 34, 34, 24, 46, 21, 43, 29, 53), "1 Kings")
    d.Add 12, ToOneBasedLongArray(Array(18, 25, 27, 44, 27, 33, 20, 29, 37, 36, 20, 22, 25, 29, 38, 20, 41, 37, 37, 21, 26, 20, 37, 20, 30), "2 Kings")
    d.Add 13, ToOneBasedLongArray(Array(54, 55, 24, 43, 26, 81, 40, 40, 44, 14, 47, 40, 14, 17, 29, 43, 27, 17, 19, 8, 30, 19, 32, 31, 31, 32, 34, 21, 30), "1 Chronicles")
    d.Add 14, ToOneBasedLongArray(Array(17, 18, 17, 22, 14, 42, 22, 18, 31, 19, 23, 16, 23, 14, 19, 14, 19, 34, 11, 37, 20, 12, 21, 27, 28, 23, 9, 27, 36, 27, 21, 33, 25, 33, 27, 23), "2 Chronicles")
    d.Add 15, ToOneBasedLongArray(Array(11, 70, 13, 24, 17, 22, 28, 36, 15, 44), "Ezra")
    d.Add 16, ToOneBasedLongArray(Array(11, 20, 32, 23, 19, 19, 73, 18, 38, 39, 36, 47, 31), "Nehemiah")
    d.Add 17, ToOneBasedLongArray(Array(22, 23, 15, 17, 14, 14, 10, 17, 32, 3), "Esther")
    d.Add 18, ToOneBasedLongArray(Array(22, 13, 26, 21, 27, 30, 21, 22, 35, 22, 20, 25, 28, 22, 35, 22, 16, 21, 29, 29, 34, 30, 17, 25, 6, 14, 23, 28, 25, 31, 40, 22, 33, 37, 16, 33, 24, 41, 30, 24, 34, 17), "Job")
    d.Add 19, ToOneBasedLongArray(Array(6, 12, 8, 8, 12, 10, 17, 9, 20, 18, 7, 8, 6, 7, 5, 11, 15, 50, 14, 9, 13, 31, 6, 10, 22, 12, 14, 9, 11, 12, 24, 11, 22, 22, 28, 12, 40, 22, 13, 17, 13, 11, 5, 26, 17, 11, 9, 14, 20, 23, 19, 9, 6, 7, 23, 13, 11, 11, 17, 12, 8, 12, 11, 10, 13, 20, 7, 35, 36, 5, 24, 20, 28, 23, 10, 12, 20, 72, 13, 19, 16, 8, 18, 12, 13, 17, 7, 18, 52, 17, 16, 15, 5, 23, 11, 13, 12, 9, 9, 5, 8, 28, 22, 35, 45, 48, 43, 13, 31, 7, 10, 10, 9, 8, 18, 19, 2, 29, 176, 7, 8, 9, 4, 8, 5, 6, 5, 6, 8, 8, 3, 18, 3, 3, 21, 26, 9, 8, 24, 13, 10, 7, 12, 15, 21, 10, 20, 14, 9, 6), "Psalms")
    d.Add 20, ToOneBasedLongArray(Array(33, 22, 35, 27, 23, 35, 27, 36, 18, 32, 31, 28, 25, 35, 33, 33, 28, 24, 29, 30, 31, 29, 35, 34, 28, 28, 27, 28, 27, 33, 31), "Proverbs")
    d.Add 21, ToOneBasedLongArray(Array(18, 26, 22, 16, 20, 12, 29, 17, 18, 20, 10, 14), "Ecclesiastes")
    d.Add 22, ToOneBasedLongArray(Array(17, 17, 11, 16, 16, 13, 13, 14), "Solomon")
    d.Add 23, ToOneBasedLongArray(Array(31, 22, 26, 6, 30, 13, 25, 22, 21, 34, 16, 6, 22, 32, 9, 14, 14, 7, 25, 6, 17, 25, 18, 23, 12, 21, 13, 29, 24, 33, 9, 20, 24, 17, 10, 22, 38, 22, 8, 31, 29, 25, 28, 28, 25, 13, 15, 22, 26, 11, 23, 15, 12, 17, 13, 12, 21, 14, 21, 22, 11, 12, 19, 12, 25, 24), "Isaiah")
    d.Add 24, ToOneBasedLongArray(Array(19, 37, 25, 31, 31, 30, 34, 22, 26, 25, 23, 17, 27, 22, 21, 21, 27, 23, 15, 18, 14, 30, 40, 10, 38, 24, 22, 17, 32, 24, 40, 44, 26, 22, 19, 32, 21, 28, 18, 16, 18, 22, 13, 30, 5, 28, 7, 47, 39, 46, 64, 34), "Jeremiah")
    d.Add 25, ToOneBasedLongArray(Array(22, 22, 66, 22, 22), "Lamentations")
    d.Add 26, ToOneBasedLongArray(Array(28, 10, 27, 17, 17, 14, 27, 18, 11, 22, 25, 28, 23, 23, 8, 63, 24, 32, 14, 49, 32, 31, 49, 27, 17, 21, 36, 26, 21, 26, 18, 32, 33, 31, 15, 38, 28, 23, 29, 49, 26, 20, 27, 31, 25, 24, 23, 35), "Ezekiel")
    d.Add 27, ToOneBasedLongArray(Array(21, 49, 30, 37, 31, 28, 28, 27, 27, 21, 45, 13), "Daniel")
    d.Add 28, ToOneBasedLongArray(Array(11, 23, 5, 19, 15, 11, 16, 14, 17, 15, 12, 14, 16, 9), "Hosea")
    d.Add 29, ToOneBasedLongArray(Array(20, 32, 21), "Joel")
    d.Add 30, ToOneBasedLongArray(Array(15, 16, 15, 13, 27, 14, 17, 14, 15), "Amos")
    d.Add 31, ToOneBasedLongArray(Array(21), "Obadiah")
    d.Add 32, ToOneBasedLongArray(Array(17, 10, 10, 11), "Jonah")
    d.Add 33, ToOneBasedLongArray(Array(16, 13, 12, 13, 15, 16, 20), "Micah")
    d.Add 34, ToOneBasedLongArray(Array(15, 13, 19), "Nahum")
    d.Add 35, ToOneBasedLongArray(Array(17, 20, 19), "Habakkuk")
    d.Add 36, ToOneBasedLongArray(Array(18, 15, 20), "Zephaniah")
    d.Add 37, ToOneBasedLongArray(Array(15, 23), "Haggai")
    d.Add 38, ToOneBasedLongArray(Array(21, 13, 10, 14, 11, 15, 14, 23, 17, 12, 17, 14, 9, 21), "Zechariah")
    d.Add 39, ToOneBasedLongArray(Array(14, 17, 18, 6), "Malachi")
    ' -------- NEW TESTAMENT --------
    d.Add 40, ToOneBasedLongArray(Array(25, 23, 17, 25, 48, 34, 29, 34, 38, 42, 30, 50, 58, 36, 39, 28, 27, 35, 30, 34, 46, 46, 39, 51, 46, 75, 66, 20), "Matthew")
    d.Add 41, ToOneBasedLongArray(Array(45, 28, 35, 41, 43, 56, 37, 38, 50, 52, 33, 44, 37, 72, 47, 20), "Mark")
    d.Add 42, ToOneBasedLongArray(Array(80, 52, 38, 44, 39, 49, 50, 56, 62, 42, 54, 59, 35, 35, 32, 31, 37, 43, 48, 47, 38, 71, 56, 53), "Luke")
    d.Add 43, ToOneBasedLongArray(Array(51, 25, 36, 54, 47, 71, 53, 59, 41, 42, 57, 50, 38, 31, 27, 33, 26, 40, 42, 31, 25), "John")
    d.Add 44, ToOneBasedLongArray(Array(26, 47, 26, 37, 42, 15, 60, 40, 43, 48, 30, 25, 52, 28, 41, 40, 34, 28, 41, 38, 40, 30, 35, 27, 27, 32, 44, 31), "Acts")
    d.Add 45, ToOneBasedLongArray(Array(32, 29, 31, 25, 21, 23, 25, 39, 33, 21, 36, 21, 14, 23, 33, 27), "Romans")
    d.Add 46, ToOneBasedLongArray(Array(31, 16, 23, 21, 13, 20, 40, 13, 27, 33, 34, 31, 13, 40, 58, 24), "1 Corinthians")
    d.Add 47, ToOneBasedLongArray(Array(24, 17, 18, 18, 21, 18, 16, 24, 15, 18, 33, 21, 14), "2 Corinthians")
    d.Add 48, ToOneBasedLongArray(Array(24, 21, 29, 31, 26, 18), "Galatians")
    d.Add 49, ToOneBasedLongArray(Array(23, 22, 21, 32, 33, 24), "Ephesians")
    d.Add 50, ToOneBasedLongArray(Array(30, 30, 21, 23), "Philippians")
    d.Add 51, ToOneBasedLongArray(Array(29, 23, 25, 18), "Colossians")
    d.Add 52, ToOneBasedLongArray(Array(10, 20, 13, 18, 28), "1 Thessalonians")
    d.Add 53, ToOneBasedLongArray(Array(12, 17, 18), "2 Thessalonians")
    d.Add 54, ToOneBasedLongArray(Array(20, 15, 16, 16, 25, 21), "1 Timothy")
    d.Add 55, ToOneBasedLongArray(Array(18, 26, 17, 22), "2 Timothy")
    d.Add 56, ToOneBasedLongArray(Array(16, 15, 15), "Titus")
    d.Add 57, ToOneBasedLongArray(Array(25), "Philemon")
    d.Add 58, ToOneBasedLongArray(Array(14, 18, 19, 16, 14, 20, 28, 13, 28, 39, 40, 29, 25), "Hebrews")
    d.Add 59, ToOneBasedLongArray(Array(27, 26, 18, 17, 20), "James")
    d.Add 60, ToOneBasedLongArray(Array(25, 25, 22, 19, 14), "1 Peter")
    d.Add 61, ToOneBasedLongArray(Array(21, 22, 18), "2 Peter")
    d.Add 62, ToOneBasedLongArray(Array(10, 29, 24, 21, 21), "1 John")
    d.Add 63, ToOneBasedLongArray(Array(13), "2 John")
    d.Add 64, ToOneBasedLongArray(Array(15), "3 John")
    d.Add 65, ToOneBasedLongArray(Array(25), "Jude")
    d.Add 66, ToOneBasedLongArray(Array(20, 29, 22, 11, 14, 17, 17, 13, 21, 11, 19, 17, 18, 20, 8, 21, 18, 24, 21, 15, 27, 21), "Revelation")

    Set GetVerseCounts = d
End Function

Public Sub GeneratePackedVerseStrings_FromDictionary()
' For each book:
'    - Each chapter's verse count is encoded as a 3-digit fixed-width number
'    - Output length = Chapters x 3
'    - Impossible to mistype
'    - Impossible to miscount
'    - Deterministic every run
' Why 3 Is Optimal
'    - Max verse count = 176 (Psalms 119)
'    - 3 digits covers up to 999
'    - Fixed width enables direct offset math
'    - Keeps implementation trivial
'    - Keeps memory small (~ 3 x 1189 chapters ~ 3567 bytes total)
'    - That is extremely compact.

    Dim d As Object
    Set d = GetVerseCounts()
    
    Dim bookID As Long
    Dim chapters As Variant
    Dim c As Long
    Dim packed As String
    
    Debug.Print "===== PACKED VERSE MAP ====="
    For bookID = 1 To 66
        If d.Exists(bookID) Then
            chapters = d(bookID)
            packed = ""
            
            For c = LBound(chapters) To UBound(chapters)
                packed = packed & Format$(chapters(c), "000")
            Next c
            
            ' Safety validation
            Debug.Assert Len(packed) = (UBound(chapters) - LBound(chapters) + 1) * 3
            Debug.Print "maps(" & bookID & ") = """ & packed & """"
        Else
            Debug.Print "Book " & bookID & " NOT FOUND"
        End If
    Next bookID
End Sub

Public Function VerifyPackedVerseMap(Optional ByVal verbose As Boolean = False) As Boolean

    Dim packedArr As Variant
    Dim expectedCounts As Variant
    
    Dim bookID As Long
    Dim chapterCount As Long
    Dim expectedChapters As Long
    Dim totalChapters As Long
    Dim totalExpected As Long
    Dim diff As Long
    
    Dim isValid As Boolean
    isValid = True
    
    ' Load packed map
    packedArr = GetChapterVerseMap()
    
    ' Load canonical expected chapter counts (should total 1189)
    expectedCounts = ExpectedChapterCounts()
    
    totalChapters = 0
    totalExpected = 0
    
    If verbose Then
        Debug.Print "Book | Chapters (Packed) | Chapters (Expected) | Diff"
        Debug.Print String(60, "=")
    End If
    
    For bookID = 1 To 66
        
        ' ----- Packed chapter count -----
        If IsArray(packedArr(bookID)) Then
            chapterCount = UBound(packedArr(bookID)) - LBound(packedArr(bookID)) + 1
        Else
            chapterCount = 0
        End If
        
        totalChapters = totalChapters + chapterCount
        
        ' ----- Canonical expected count (Array() is 0-based) -----
        expectedChapters = expectedCounts(bookID - 1)
        totalExpected = totalExpected + expectedChapters
        
        diff = chapterCount - expectedChapters
        
        If diff <> 0 Then
            isValid = False
        End If
        
        If verbose Then
            Debug.Print "Book " & bookID & _
                        " | " & chapterCount & _
                        " | " & expectedChapters & _
                        " | " & diff
        End If
                    
    Next bookID
    
    If totalChapters <> totalExpected Then
        isValid = False
    End If
    
    If verbose Then
        Debug.Print String(60, "-")
        Debug.Print "EXPECTED Total Chapters: " & totalExpected
        Debug.Print "ACTUAL Total Chapters:   " & totalChapters
        Debug.Print "DIFFERENCE:              " & (totalChapters - totalExpected)
    End If
    
    VerifyPackedVerseMap = isValid

End Function



