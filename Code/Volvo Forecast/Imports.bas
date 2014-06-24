Attribute VB_Name = "Imports"
Option Explicit

Sub ImportData()
    Dim iRows As Long
    Dim iCols As Integer
    Dim aHeaders As Variant

    Sheets("Drop In").Select

    UserImportFile Sheets("Drop In").Range("A1"), True
    iRows = ActiveSheet.UsedRange.Rows.Count + 1

    UserImportFile ActiveSheet.Cells(iRows, 1), True
    iRows = ActiveSheet.UsedRange.Rows.Count + 1

    UserImportFile ActiveSheet.Cells(iRows, 1), True
    iRows = ActiveSheet.UsedRange.Rows.Count
    iCols = ActiveSheet.UsedRange.Columns.Count

    aHeaders = Range(Cells(1, 1), Cells(1, iCols)).Value

    Range("A1").AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(iRows, iCols)).AutoFilter Field:=1, Criteria1:="PLNTCODE"
    Cells.Delete Shift:=xlUp
    Rows(1).Insert

    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))).Value = aHeaders
    Columns("S:S").Delete
End Sub

Sub ImportExpediteNotes()
    Dim sPath As String
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Integer

    For i = 1 To 14
        sPath = "\\br3615gaps\gaps\Volvo\" & Format(Date, "yyyy") & " Alerts\Slink Alert " & Format(Date - i, "m-dd-yy") & ".xlsx"
        If FileExists(sPath) Then
            Workbooks.Open sPath
            Sheets("Expedite").Select
            TotalRows = ActiveSheet.UsedRange.Rows.Count
            TotalCols = ActiveSheet.UsedRange.Columns.Count

            Range(Cells(1, 1), Cells(TotalRows, 1)).Copy Destination:=ThisWorkbook.Sheets("Expedite Notes").Range("A1")
            Range(Cells(1, TotalCols), Cells(TotalRows, TotalCols)).Copy Destination:=ThisWorkbook.Sheets("Expedite Notes").Range("B1")

            ActiveWorkbook.Close

            Sheets("Forecast").Select
            TotalRows = ActiveSheet.UsedRange.Rows.Count
            TotalCols = ActiveSheet.UsedRange.Columns.Count

            Cells(2, TotalCols).Formula = "=IFERROR(IF(VLOOKUP(A2,'Expedite Notes'!A:B,2,FALSE)=0,"""",VLOOKUP(A2,'Expedite Notes'!A:B,2,FALSE)),"""")"
            Cells(2, TotalCols).AutoFill Destination:=Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))

            With Range(Cells(2, TotalCols), Cells(TotalRows, TotalCols))
                .Value = .Value
            End With

            Exit For
        End If
    Next
End Sub

Sub ImportMaster()
    Const Path As String = "\\br3615gaps\gaps\Billy Mac-Master Lists\"
    Dim Wkbk As Workbook
    Dim File As String

    File = "Volvo Master List " & Format(Date, "yyyy") & ".xlsx"

    Workbooks.Open Path & File
    Set Wkbk = ActiveWorkbook
    Sheets("ACTIVE").Select
    ActiveSheet.UsedRange.Copy

    ThisWorkbook.Activate
    Sheets("Master").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues, _
                             Operation:=xlNone, _
                             SkipBlanks:=False, _
                             Transpose:=False
    Application.CutCopyMode = False
    Wkbk.Close
    
    Sheets("Macro").Select
End Sub
