Attribute VB_Name = "basSectionPageSetup"
Option Explicit
Option Compare Text

' Classification: shared helper for the scripture-only export (Pass 5) and the RWB front matter (basRwbFrontMatter). Not used by runtime ribbon navigation.
'
' basSectionPageSetup - copy and compare the page setup of two sections. A Word section break gives the new section the page setup of the
' section it splits (the new RWB front matter took the Study Bible title page's US Letter layout, FU-34, 2026-11-11), so the page setup
' that the publication wants is copied explicitly from a body section, and the export checks that it is the same.

Private Const SIZE_TOLERANCE_POINTS As Double = 0.5

' Copies the page-setup properties that matter from one Section to another. Headers and footers are NOT copied (plan task 4).
Public Sub CopySectionPageSetup(ByVal src As Object, ByVal dst As Object)
    Dim ps As Object
    Dim pd As Object
    Set ps = src.PageSetup
    Set pd = dst.PageSetup
    pd.Orientation = ps.Orientation
    pd.pageWidth = ps.pageWidth
    pd.PageHeight = ps.PageHeight
    pd.TopMargin = ps.TopMargin
    pd.BottomMargin = ps.BottomMargin
    pd.leftMargin = ps.leftMargin
    pd.rightMargin = ps.rightMargin
    pd.gutter = ps.gutter
    pd.HeaderDistance = ps.HeaderDistance
    pd.FooterDistance = ps.FooterDistance
    pd.SectionStart = ps.SectionStart
    pd.DifferentFirstPageHeaderFooter = ps.DifferentFirstPageHeaderFooter
    pd.OddAndEvenPagesHeaderFooter = ps.OddAndEvenPagesHeaderFooter
    pd.TextColumns.SetCount ps.TextColumns.Count
    pd.TextColumns.LineBetween = ps.TextColumns.LineBetween
    ' Word refuses to set the width or the spacing of a single column (error 5993, 2026-11-11).
    If ps.TextColumns.Count > 1 Then
        pd.TextColumns.EvenlySpaced = ps.TextColumns.EvenlySpaced
        If ps.TextColumns.EvenlySpaced Then pd.TextColumns.Spacing = ps.TextColumns.Spacing
    End If
End Sub

' True when the two sections have the same page size, orientation, margins, gutter and header/footer distances (within half a point).
Public Function SamePageLayout(ByVal a As Object, ByVal b As Object) As Boolean
    Dim pa As Object
    Dim pb As Object
    Set pa = a.PageSetup
    Set pb = b.PageSetup
    SamePageLayout = pa.Orientation = pb.Orientation _
        And Near(pa.pageWidth, pb.pageWidth) And Near(pa.PageHeight, pb.PageHeight) _
        And Near(pa.TopMargin, pb.TopMargin) And Near(pa.BottomMargin, pb.BottomMargin) _
        And Near(pa.leftMargin, pb.leftMargin) And Near(pa.rightMargin, pb.rightMargin) _
        And Near(pa.gutter, pb.gutter) _
        And Near(pa.HeaderDistance, pb.HeaderDistance) And Near(pa.FooterDistance, pb.FooterDistance)
End Function

Private Function Near(ByVal x As Double, ByVal y As Double) As Boolean
    Near = Abs(x - y) <= SIZE_TOLERANCE_POINTS
End Function
