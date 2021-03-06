Attribute VB_Name = "Program"
Option Explicit

Public Const VersionNumber As String = "1.0.0"
Public Const RepositoryName As String = "Volvo_Forecast"

Sub Main()
    Dim Branch As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ERROR
    Branch = InputBox("Enter your branch number", "Branch Entry")
    If Branch = "" Then Exit Sub
    ImportGaps SimsAsText:=False, Branch:=Branch
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
    Email SendTo:="jogardner@wesco.com; jlquatra@wesco.com; bford@wesco.com", _
          Subject:="Volvo Forecast", _
          Body:="A new forecast is available on the network <a href=""\\br3615gaps\gaps\Volvo\" & _
                Format(Date, "yyyy") & " Alerts\Slink Alert " & Format(Date, "m-dd-yy") & ".xlsx""" & ">here</a>."
    Clean
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
