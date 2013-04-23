Attribute VB_Name = "Exports"
Option Explicit

Sub ExportForecast()
    Sheets("Forecast").Copy
    ThisWorkbook.Sheets("Non-Stock").Copy After:=ActiveWorkbook.Sheets(1)
    ThisWorkbook.Sheets("Master").Copy After:=ActiveWorkbook.Sheets(2)
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(3)
    ActiveSheet.Name = "Expedite"
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(4)
    ActiveSheet.Name = "Order"
    
    ActiveWorkbook.Sheets(1).Select
    ActiveWorkbook.SaveAs "\\br3615gaps\gaps\Volvo\2013 Alerts\Slink Alert " & Format(Date, "m-dd-yy") & ".xlsx", xlOpenXMLWorkbook
    ActiveWorkbook.Close

    Sheets("Temp").Copy
    ActiveWorkbook.SaveAs "\\br3615gaps\gaps\Volvo\2013 Slink\Combined " & Format(Date, "m-dd-yy") & ".xlsx", xlOpenXMLWorkbook
    ActiveWorkbook.Close
End Sub

