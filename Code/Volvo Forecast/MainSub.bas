Attribute VB_Name = "MainSub"
Option Explicit

Sub Main()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ERROR
    ImportGaps
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
          Subject:="Volvo Forecast", _
          Body:="""\\br3615gaps\gaps\Volvo\2013 Alerts\Slink Alert " & Format(Date, "m-dd-yy") & ".xlsx"""
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ERROR:
    Exit Sub

End Sub

Sub Clean()
    Dim s As Variant

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Master" And s.Name <> "Macro" Then
            s.Cells.Delete
        End If
    Next
End Sub
