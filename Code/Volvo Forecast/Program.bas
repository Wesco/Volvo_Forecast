Attribute VB_Name = "Program"
Option Explicit

Sub Main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ERROR
    ImportGaps SimsAsText:=False
    ImportMaster
    ImportData
    On Error GoTo 0
    CombineData
    SeparateNonStock
    CreatePivotTable
    CreateForecast
    ImportExpediteNotes
    ExportForecast
    MsgBox ("Complete!")
    Email SendTo:="JBarnhill@wesco.com", _
          CC:="ACoffey@wesco.com", _
          Subject:="Volvo Forecast", _
          Body:="""\\br3615gaps\gaps\Volvo\" & Format(Date, "yyyy") & " Alerts\Slink Alert " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ERROR:
    Clean
    Exit Sub

End Sub

Sub Clean()
    Dim s As Variant

    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C8").Select
End Sub
