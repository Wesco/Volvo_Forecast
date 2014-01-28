Attribute VB_Name = "Exports"
Option Explicit

Sub ExportForecast()
    Dim FilePath As String
    Dim FileName As String

    Sheets("Forecast").Copy
    ThisWorkbook.Sheets("Non-Stock").Copy After:=ActiveWorkbook.Sheets(1)
    ThisWorkbook.Sheets("Master").Copy After:=ActiveWorkbook.Sheets(2)
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(3)
    ActiveSheet.Name = "Expedite"
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(4)
    ActiveSheet.Name = "Order"

    FilePath = "\\br3615gaps\gaps\Volvo\" & Format(Date, "yyyy") & " Alerts\"
    FileName = "Slink Alert " & Format(Date, "m-dd-yy") & ".xlsx"
    If Not FolderExists(FilePath) Then RecMkDir FilePath

    ActiveWorkbook.Sheets(1).Select
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    ActiveWorkbook.Close

    FilePath = "\\br3615gaps\gaps\Volvo\" & Format(Date, "yyyy") & " Slink\"
    FileName = "Combined " & Format(Date, "m-dd-yy") & ".xlsx"
    If Not FolderExists(FilePath) Then RecMkDir FilePath

    Sheets("Temp").Copy
    ActiveWorkbook.SaveAs FilePath & FileName, xlOpenXMLWorkbook
    ActiveWorkbook.Close
End Sub

