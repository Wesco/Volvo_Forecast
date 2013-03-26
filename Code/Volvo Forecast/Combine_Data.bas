Attribute VB_Name = "Combine_Data"
Option Explicit

Sub CombineData()
    Dim iRows As Long

    Sheets("Drop In").Select
    iRows = ActiveSheet.UsedRange.Rows.Count

    Columns("A:A").Insert
    Range("A1").Value = "SIM"
    Range("A2").Formula = "=IF(IFERROR(VLOOKUP(G2,Master!A:B,2,FALSE),"""")=0,"""",IFERROR(VLOOKUP(G2,Master!A:B,2,FALSE),""""))"
    Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
    With Range(Cells(2, 1), Cells(iRows, 1))
        .Value = .Value
    End With

    ActiveSheet.UsedRange.Replace What:="=""", _
                                  Replacement:="", _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  MatchCase:=False, _
                                  SearchFormat:=False, _
                                  ReplaceFormat:=False

    ActiveSheet.UsedRange.Replace What:="""", _
                                  Replacement:="", _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  MatchCase:=False, _
                                  SearchFormat:=False, _
                                  ReplaceFormat:=False

    Columns("F:F").TextToColumns Destination:=Range("F1"), _
                                 DataType:=xlDelimited, _
                                 FieldInfo:=Array(1, 5)

    iRows = ActiveSheet.UsedRange.Rows.Count

    Columns("F:F").Insert Shift:=xlToRight
    Range("F1").Value = "DUEDT"
    Range("F2").Formula = "=TEXT(G2,""mmm-yyyy"")"
    Range("F2").AutoFill Destination:=Range(Cells(2, 6), Cells(iRows, 6))
    With Range(Cells(2, 6), Cells(iRows, 6))
        .Value = .Value
    End With
    Columns("G:G").Delete
End Sub

Sub CreatePivotTable()
    Dim iRows As Long
    Dim iCols As Integer

    Sheets("Drop In").Select

    iRows = ActiveSheet.UsedRange.Rows.Count
    iCols = ActiveSheet.UsedRange.Columns.Count

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                      SourceData:=Sheets("Drop In").Range(Cells(1, 1), Cells(iRows, 18)), _
                                      Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:="PivotTable!R1C1", _
                                      TableName:="PivotTable1", _
                                      DefaultVersion:=xlPivotTableVersion14

    Sheets("PivotTable").Select
    Cells(1, 1).Select
    
    'Combine by SIM/Part
    '    With ActiveSheet.PivotTables("PivotTable1")
    '        .PivotFields("SIM").Orientation = xlRowField
    '        .PivotFields("SIM").LayoutForm = xlTabular
    '        .PivotFields("SIM").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    '        .PivotFields("SIM").Position = 1
    '        .PivotFields("EDCSPT").Orientation = xlRowField
    '        .PivotFields("EDCSPT").LayoutForm = xlTabular
    '        .PivotFields("EDCSPT").Position = 2
    '        .AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("QTYDU"), "Sum of QTYDU", xlSum
    '        .PivotFields("DUEDT").Orientation = xlColumnField
    '        .PivotFields("DUEDT").Position = 1
    '    End With

    'Combine by Part
    '    With ActiveSheet.PivotTables("PivotTable1")
    '        .PivotFields("EDCSPT").Orientation = xlRowField
    '        .PivotFields("EDCSPT").Position = 1
    '        .AddDataField .PivotFields("QTYDU"), "Sum of QTYDU", xlSum
    '        .PivotFields("DUEDT").Orientation = xlColumnField
    '        .PivotFields("DUEDT").Position = 1
    '    End With

    'Combine by SIM
    With ActiveSheet.PivotTables("PivotTable1")
        .PivotFields("SIM").Orientation = xlRowField
        .PivotFields("SIM").Position = 1
        .PivotFields("DUEDT").Orientation = xlColumnField
        .PivotFields("DUEDT").Position = 1
        .AddDataField .PivotFields("QTYDU"), "Sum of QTYDU", xlSum
    End With

    Cells.Copy
    Cells.PasteSpecial Paste:=xlPasteValues, _
                       Operation:=xlNone, _
                       SkipBlanks:=False, _
                       Transpose:=False
    Application.CutCopyMode = False
    ActiveSheet.UsedRange.Replace What:="(blank)", _
                                  Replacement:="", _
                                  LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, _
                                  MatchCase:=False, _
                                  SearchFormat:=False, _
                                  ReplaceFormat:=False

    ActiveSheet.UsedRange.Copy Destination:=Sheets("Temp").Range("A1")
    Sheets("Temp").Select

    Rows(1).Delete
    Range("P:ZZ").Delete
    Columns(ActiveSheet.UsedRange.Columns.Count).Delete
    Rows(ActiveSheet.UsedRange.Rows.Count).Delete

    iRows = ActiveSheet.UsedRange.Rows.Count
    iCols = ActiveSheet.UsedRange.Columns.Count

    Range(Cells(1, 2), Cells(iRows, iCols)).SpecialCells(xlCellTypeBlanks).Value = 0
    Range(Cells(1, 1), Cells(1, iCols)).NumberFormat = "mmm-yyy"
    Columns("A:A").Insert
    Range("A1").Value = "Part"
    Range("B1").Value = "SIM"
    Range("A2").Formula = "=IFERROR(IF(VLOOKUP(B2,'Drop In'!A:G,7,FALSE)=0,"""",VLOOKUP(B2,'Drop In'!A:G,7,FALSE)),"""")"
    Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
    Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value
    Columns("B:B").Cut
    Columns("A:A").Insert Shift:=xlToRight
    Application.CutCopyMode = False
End Sub




