Attribute VB_Name = "basWordRepairRunner"
Option Explicit
Option Compare Text
Option Private Module

Public Const MODULE_NOT_EMPTY_DUMMY As String = vbNullString

Public Sub RunRepairWrappedVerseMarkers_Across_Pages_From(startPage As Long)
    Dim totalFixes As Long, pgFixCount As Long
    Dim numPages As Long: numPages = 0 ' Adjust if scanning more than one page

    Dim sessionID As String
    sessionID = Format(Now, "yyyyMMdd_HHmmss")

    Dim logPath As String
    logPath = "C:\adaept\aeBibleClass\rpt\RepairLog.txt"

    Dim logFile As Integer
    logFile = FreeFile

    ' Create file with header if it doesn't exist
    If Dir(logPath) = "" Then
        Open logPath For Output As #logFile
        Print #logFile, "SessionID,PageNum,Repairs"
        Close #logFile
    End If

    ' Append results
    Open logPath For Append As #logFile
    Dim p As Long
    For p = startPage To startPage + numPages
        pgFixCount = 0
        RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage p, pgFixCount
        Print #logFile, sessionID & "," & p & "," & pgFixCount
        totalFixes = totalFixes + pgFixCount
    Next p
    Close #logFile

    'MsgBox "Repair complete. CSV log updated at:" & vbCrLf & logPath, vbInformation
    Selection.GoTo What:=wdGoToPage, name:=CStr(startPage)
End Sub

Public Sub RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext_SinglePage(pageNum As Long, ByRef fixCount As Long)
    ' Same logic as full macro, but suppresses MsgBox and passes fixCount by reference.
    ' Copy the full body from RepairWrappedVerseMarkers_MergedPrefix_ByColumnContext here
    ' And replace `MsgBox` line with: fixCount = fixCount
    Dim pgRange As range, ch As range, scanRange As range, prefixCh As range
    Dim pageStart As Long, pageEnd As Long
    Dim chapterMarker As String, verseDigits As String, combinedNumber As String
    Dim markerStart As Long, markerEnd As Long, verseEnd As Long
    Dim prefixTxt As String, prefixStyle As String, prefixAsc As Variant
    Dim prefixY As Single, digitY As Single, digitX As Single
    Dim nextWords As String, lookAhead As range, token As range, wCount As Integer
    Dim logBuffer As String
    Dim ascii12Count As Long
    Dim ascii160MissingCount As Long
    Dim suffix160Count As Long
    Dim suffixHairSpaceCount As Long
    Dim suffixSpaceCount As Long
    Dim suffixOtherCount As Long

    fixCount = 0
    logBuffer = "=== Smart Prefix Repair on Page " & pageNum & " ===" & vbCrLf

    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum))
    pageStart = pgRange.Start
    Set pgRange = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pageNum + 1))
    pageEnd = pgRange.Start - 1

    Dim i As Long
    i = pageStart
    Dim headerText As String
    headerText = GetPageHeaderText(pageNum)
    'Debug.Print "Page " & pageNum & " header: " & headerText
    logBuffer = logBuffer & "Header for page " & pageNum & ": " & headerText & vbCrLf
    
    Do While i < pageEnd
        Set ch = ActiveDocument.range(i, i + 1)
        If Len(Trim(ch.text)) = 1 And IsNumeric(ch.text) And ch.style.NameLocal = "Chapter Verse marker" And ch.font.color = RGB(255, 165, 0) Then
            ' Assemble chapter marker block
            chapterMarker = ch.text
            markerStart = i
            markerEnd = i + 1
            Do While markerEnd < pageEnd
                Set scanRange = ActiveDocument.range(markerEnd, markerEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Chapter Verse marker" And scanRange.font.color = RGB(255, 165, 0) Then
                        chapterMarker = chapterMarker & scanRange.text
                        markerEnd = markerEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            digitY = ch.Information(wdVerticalPositionRelativeToPage)
            digitX = ch.Information(wdHorizontalPositionRelativeToPage)

            ' Assemble verse marker block
            verseDigits = ""
            verseEnd = markerEnd
            Do While verseEnd < pageEnd
                Set scanRange = ActiveDocument.range(verseEnd, verseEnd + 1)
                If Len(Trim(scanRange.text)) = 1 And IsNumeric(scanRange.text) Then
                    If scanRange.style.NameLocal = "Verse marker" And scanRange.font.color = RGB(80, 200, 120) Then
                        verseDigits = verseDigits & scanRange.text
                        verseEnd = verseEnd + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop

            If Len(verseDigits) > 0 Then
                combinedNumber = chapterMarker & verseDigits
            
                Dim chInfo As range
                Set chInfo = ActiveDocument.range(verseEnd, verseEnd + 1)
                'Debug.Print "Hair space font: " & chInfo.font.name & " | Size=" & chInfo.font.Size & " | Style=" & chInfo.style.NameLocal & " | ASCII=" & AscW(chInfo.text)
                
                Dim suffixCh As range
                Set suffixCh = ActiveDocument.range(verseEnd, verseEnd + 1)
                Dim suffixAsc As Long
                suffixAsc = AscW(suffixCh.text)

                Select Case suffixAsc
                    Case 160: suffix160Count = suffix160Count + 1
                    Case 8239: suffixHairSpaceCount = suffixHairSpaceCount + 1
                    Case 32: suffixSpaceCount = suffixSpaceCount + 1
                    Case Else: suffixOtherCount = suffixOtherCount + 1
                End Select

                ' Optional diagnostic
                'Debug.Print "Suffix [" & combinedNumber & "] ASCII=" & suffixAsc & " Style=" & suffixCh.style.NameLocal & " Font=" & suffixCh.font.name & " Size=" & suffixCh.font.Size
                
                ' Chr(12) audit
                If Len(combinedNumber) = 1 And AscW(combinedNumber) = 12 Then
                    ascii12Count = ascii12Count + 1
                    i = verseEnd
                    GoTo SkipLogging
                End If
                
                ' Prefix check
                If markerStart > pageStart Then
                    Set prefixCh = ActiveDocument.range(markerStart - 1, markerStart)
                    prefixTxt = prefixCh.text
                    prefixStyle = prefixCh.style.NameLocal
                    prefixAsc = AscW(prefixTxt)
                    Debug.Print headerText & " " & chapterMarker & ":" & verseDigits, prefixAsc    ', combinedNumber

                    prefixY = prefixCh.Information(wdVerticalPositionRelativeToPage)

                    If (prefixAsc = 32 Or prefixAsc = 160) And prefixStyle = "Normal" Then
                        If Abs(prefixY - digitY) < 25 Then
                            nextWords = ""
                            Set lookAhead = ActiveDocument.range(verseEnd, verseEnd + 80)
                            wCount = 0
                            For Each token In lookAhead.words
                                If token.text Like "*^13*" Then Exit For
                                If Trim(token.text) <> "" Then
                                    nextWords = nextWords & Trim(token.text) & " "
                                    wCount = wCount + 1
                                    If wCount = 2 Then Exit For
                                End If
                            Next token

                            ' Column edge logic
                            If digitX < 50 Then
                                prefixCh.text = vbCr
                                logBuffer = logBuffer & "> Repaired prefix before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | Break inserted | Next words:  " & Trim(nextWords) & " " & vbCrLf
                            Else
                                prefixCh.text = ""
                                logBuffer = logBuffer & "> Removed space before '" & combinedNumber & "' @ X=" & Format(digitX, "0.0") & " | No break | Next words:  " & Trim(nextWords) & " " & vbCrLf
                            End If

                            fixCount = fixCount + 1
                        End If
                    End If
                'End If
                ElseIf markerStart = pageStart Then
                    logBuffer = logBuffer & "Marker '" & combinedNumber & "' is at the very start of page " & pageNum & vbCrLf
                    Debug.Print headerText & " " & chapterMarker & ":" & verseDigits, "SoP"    ', combinedNumber
                End If

                i = verseEnd
            Else
                i = markerEnd
            End If
        Else
            i = i + 1
        End If
SkipLogging:
    Loop

    logBuffer = logBuffer & "=== " & fixCount & " markers repaired on page " & pageNum & " ==="
    logBuffer = logBuffer & vbCrLf & "ASCII 12 audit: " & ascii12Count & " marker(s) on page " & pageNum & " contain Chr(12)"
    logBuffer = logBuffer & vbCrLf & "ASCII 160 audit: " & ascii160MissingCount & " marker(s) on page " & pageNum & " missing Chr(160) suffix" & vbCrLf
    Debug.Print logBuffer
    'MsgBox fixCount & " marker(s) repaired on page " & pageNum & ".", vbInformation
    fixCount = fixCount
    Selection.GoTo What:=wdGoToPage, name:=CStr(pageNum)
End Sub

Public Function GetPageHeaderText(pgNum As Long) As String
    Dim rng As range
    Dim sec As section
    Dim hdr As HeaderFooter
    
    ' Get range for the page
    Set rng = ActiveDocument.GoTo(What:=wdGoToPage, name:=CStr(pgNum))
    Set sec = rng.Sections(1)   ' Page belongs to exactly one Section
    
    ' Default to primary header
    Set hdr = sec.Headers(wdHeaderFooterPrimary)
    
    ' If primary is empty, check for first-page or even-page headers
    'If Len(hdr.range.text) = 0 Then
    '    If sec.Headers(wdHeaderFooterFirstPage).Exists Then
    '        Set hdr = sec.Headers(wdHeaderFooterFirstPage)
    '    ElseIf sec.Headers(wdHeaderFooterEvenPages).Exists Then
    '        Set hdr = sec.Headers(wdHeaderFooterEvenPages)
    '    End If
    'End If
    
    ' Clean up the header text (Word stores an end-of-cell marker)
    GetPageHeaderText = TitleCase(Trim(Replace(hdr.range.text, Chr(13), " ")))
End Function

Public Function TitleCase(ByVal txt As String) As String
    Dim words() As String
    Dim i As Integer

    ' Split the sentence into words
    words = Split(LCase(txt), " ")

    ' Capitalize each word
    For i = 0 To UBound(words)
        If Len(words(i)) > 0 Then
            words(i) = UCase(Left(words(i), 1)) & mid(words(i), 2)
        End If
    Next i

    ' Recombine the words into a sentence
    TitleCase = Join(words, " ")
End Function

