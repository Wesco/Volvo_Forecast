Attribute VB_Name = "BuildForecast"
Option Explicit

Sub CreateForecast()
    Dim iRows As Long
    Dim iCols As Long
    Dim aHeaders As Variant
    Dim aMonths As Variant

    Sheets("Temp").Select
    iRows = ActiveSheet.UsedRange.Rows.Count
    Range(Cells(1, 1), Cells(iRows, 2)).Copy Destination:=Sheets("Forecast").Range("A1")

    Sheets("Forecast").Select
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Headers
    aHeaders = Array("Description", _
                     "OH", _
                     "OR", _
                     "OO", _
                     "BO", _
                     "WDC", _
                     "Last Cost", _
                     "UOM", _
                     "Supplier")
    Range("C1:K1").Value = aHeaders

    'Description
    [C2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:F,6,FALSE),"""")"
    [C2].AutoFill Destination:=Range(Cells(2, 3), Cells(iRows, 3))

    'On Hand
    [D2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:G,7,FALSE),0)"
    [D2].AutoFill Destination:=Range(Cells(2, 4), Cells(iRows, 4))

    'Reserve
    [E2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:H,8,FALSE),0)"
    [E2].AutoFill Destination:=Range(Cells(2, 5), Cells(iRows, 5))

    'OO
    [F2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:J,10,FALSE),0)"
    [F2].AutoFill Destination:=Range(Cells(2, 6), Cells(iRows, 6))

    'BO
    [G2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:I,9,FALSE),0)"
    [G2].AutoFill Destination:=Range(Cells(2, 7), Cells(iRows, 7))

    'WDC
    [H2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:AK,37,FALSE),0)"
    [H2].AutoFill Destination:=Range(Cells(2, 8), Cells(iRows, 8))

    'Last Cost
    [I2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:AF,32,FALSE),0)"
    [I2].AutoFill Destination:=Range(Cells(2, 9), Cells(iRows, 9))

    'UOM
    [J2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:AJ,36,FALSE),"""")"
    [J2].AutoFill Destination:=Range(Cells(2, 10), Cells(iRows, 10))

    'Supplier
    [K2].Formula = "=IFERROR(VLOOKUP(A2,Gaps!A:AM,39,FALSE),"""")"
    [K2].AutoFill Destination:=Range(Cells(2, 11), Cells(iRows, 11))



    'Months
    aMonths = Array("=Temp!C1", _
                    "=Temp!D1", _
                    "=Temp!E1", _
                    "=Temp!F1", _
                    "=Temp!G1", _
                    "=Temp!H1", _
                    "=Temp!I1", _
                    "=Temp!J1", _
                    "=Temp!K1", _
                    "=Temp!L1", _
                    "=Temp!M1", _
                    "=Temp!N1")
    Range("L1:W1").Formula = aMonths
    Range("L1:W1").NumberFormat = "mmm-yy"

    [L2].Formula = "=D2-VLOOKUP(B2,Temp!B:C,2,FALSE)"
    [L2].AutoFill Destination:=Range(Cells(2, 12), Cells(iRows, 12))

    [M2].Formula = "=L2-VLOOKUP(B2,Temp!B:D,3,FALSE)"
    [M2].AutoFill Destination:=Range(Cells(2, 13), Cells(iRows, 13))

    [N2].Formula = "=M2-VLOOKUP(B2,Temp!B:E,4,FALSE)"
    [N2].AutoFill Destination:=Range(Cells(2, 14), Cells(iRows, 14))

    [O2].Formula = "=N2-VLOOKUP(B2,Temp!B:F,5,FALSE)"
    [O2].AutoFill Destination:=Range(Cells(2, 15), Cells(iRows, 15))

    [P2].Formula = "=O2-VLOOKUP(B2,Temp!B:G,6,FALSE)"
    [P2].AutoFill Destination:=Range(Cells(2, 16), Cells(iRows, 16))

    [Q2].Formula = "=P2-VLOOKUP(B2,Temp!B:H,7,FALSE)"
    [Q2].AutoFill Destination:=Range(Cells(2, 17), Cells(iRows, 17))

    [R2].Formula = "=Q2-VLOOKUP(B2,Temp!B:I,8,FALSE)"
    [R2].AutoFill Destination:=Range(Cells(2, 18), Cells(iRows, 18))

    [S2].Formula = "=R2-VLOOKUP(B2,Temp!B:J,9,FALSE)"
    [S2].AutoFill Destination:=Range(Cells(2, 19), Cells(iRows, 19))

    [T2].Formula = "=S2-VLOOKUP(B2,Temp!B:K,10,FALSE)"
    [T2].AutoFill Destination:=Range(Cells(2, 20), Cells(iRows, 20))

    [U2].Formula = "=T2-VLOOKUP(B2,Temp!B:L,11,FALSE)"
    [U2].AutoFill Destination:=Range(Cells(2, 21), Cells(iRows, 21))

    [V2].Formula = "=U2-VLOOKUP(B2,Temp!B:M,12,FALSE)"
    [V2].AutoFill Destination:=Range(Cells(2, 22), Cells(iRows, 22))

    [W2].Formula = "=V2-VLOOKUP(B2,Temp!B:N,13,FALSE)"
    [W2].AutoFill Destination:=Range(Cells(2, 23), Cells(iRows, 23))

    [X1].Value = "Notes"
    [X2].Formula = "=IF(IFERROR(VLOOKUP(B2,Master!A:Q,17,FALSE),"""")=0,"""",IFERROR(VLOOKUP(B2,Master!A:Q,17,FALSE),""""))"
    [X2].AutoFill Destination:=Range(Cells(2, 24), Cells(iRows, 24))
    [Y1].Value = "Expedite Notes"

    ActiveSheet.UsedRange.Value = ActiveSheet.UsedRange.Value
    ActiveSheet.UsedRange.Cells.HorizontalAlignment = xlLeft
    Columns("Z:ZZ").Delete

    iRows = ActiveSheet.UsedRange.Rows.Count
    iCols = ActiveSheet.UsedRange.Columns.Count

    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(iRows, iCols)), , xlYes).Name = "Table1"

    Range("Table1[#All]").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("L2").Select
    Range("L2").SparklineGroups.Add Type:=xlSparkColumn, SourceData:="M2:X2"
    Range("L1").Value = "Summary"

    With Selection.SparklineGroups.Item(1)
        .SeriesColor.ThemeColor = 5
        .SeriesColor.TintAndShade = -0.499984740745262
        .Points.Negative.Color.ThemeColor = 6
        .Points.Negative.Color.TintAndShade = 0
        .Points.Markers.Color.ThemeColor = 5
        .Points.Markers.Color.TintAndShade = -0.499984740745262
        .Points.Highpoint.Color.ThemeColor = 5
        .Points.Highpoint.Color.TintAndShade = 0
        .Points.Lowpoint.Color.ThemeColor = 5
        .Points.Lowpoint.Color.TintAndShade = 0
        .Points.Firstpoint.Color.ThemeColor = 5
        .Points.Firstpoint.Color.TintAndShade = 0.399975585192419
        .Points.Lastpoint.Color.ThemeColor = 5
        .Points.Lastpoint.Color.TintAndShade = 0.399975585192419
        .SeriesColor.Color = 3289650
        .SeriesColor.TintAndShade = 0
        .Points.Negative.Color.Color = 208
        .Points.Negative.Color.TintAndShade = 0
        .Points.Markers.Color.Color = 208
        .Points.Markers.Color.TintAndShade = 0
        .Points.Highpoint.Color.Color = 208
        .Points.Highpoint.Color.TintAndShade = 0
        .Points.Lowpoint.Color.Color = 208
        .Points.Lowpoint.Color.TintAndShade = 0
        .Points.Firstpoint.Color.Color = 208
        .Points.Firstpoint.Color.TintAndShade = 0
        .Points.Lastpoint.Color.Color = 208
        .Points.Lastpoint.Color.TintAndShade = 0
        .Points.Negative.Visible = True
    End With

    Selection.AutoFill Destination:=Range("Table1[Summary]"), Type:=xlFillDefault
    Range("Table1[Summary]").Select
    Columns("L:L").EntireColumn.AutoFit
    Range("Table1[[#Headers],[Summary]]").Select

    Range("L2").Select
    ActiveSheet.UsedRange.Columns.AutoFit
    Range("A1").Select

    ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort.SortFields.Add Key:= _
                                                                                    Range("Table1[[#All],[SIM]]"), _
                                                                                    SortOn:=xlSortOnValues, _
                                                                                    Order:=xlAscending, _
                                                                                    DataOption:=xlSortNormal

    With ActiveWorkbook.Worksheets("Forecast").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Columns("L:N").Insert
    [L1].Value = "LT/Days"
    [L2].Formula = "=IFERROR(VLOOKUP(B2,Master!A:J,10,FALSE),0)"

    [M1].Value = "LT/Weeks"
    [M2].Formula = "=IFERROR(VLOOKUP(B2,Master!A:J,10,FALSE),0)/7"

    [N1].Value = "Min Qty"
    [N2].Formula = "=IFERROR(VLOOKUP(B2,Master!A:K,11,FALSE),0)"

    Range(Cells(1, 12), Cells(iRows, 14)).Value = Range(Cells(1, 12), Cells(iRows, 14)).Value
    
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub



