Attribute VB_Name = "SeparateData"
Option Explicit

Sub SeparateNonStock()
    Dim iCols As Long
    Dim iRows As Long
    Dim aHeaders As Variant

    Sheets("Drop In").Select
    iCols = ActiveSheet.UsedRange.Columns.Count
    iRows = ActiveSheet.UsedRange.Rows.Count
    aHeaders = Range(Cells(1, 1), Cells(1, iCols))

    Range("A1").AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(iRows, iCols)).AutoFilter Field:=1, Criteria1:="="
    ActiveSheet.UsedRange.Copy Destination:=Sheets("Non-Stock").Range("A1")
    ActiveSheet.UsedRange.Delete
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(Cells(1, 1), Cells(1, iCols)) = aHeaders

    Sheets("Non-Stock").Columns("A:A").Delete
End Sub

