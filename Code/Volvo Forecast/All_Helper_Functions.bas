Attribute VB_Name = "All_Helper_Functions"
Option Explicit
'Pauses for x# of milliseconds
'Used for email function to prevent
'all emails from being sent at once
'Example: "Sleep 1500" will pause for 1.5 seconds
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Proc  : Function FileExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a file exists
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Function FileExists(ByVal sPath As String) As Boolean
    'Remove trailing backslash
    If InStr(Len(sPath), sPath, "\") > 0 Then sPath = Left(sPath, Len(sPath) - 1)
    'Check to see if the directory exists and return true/false
    If Dir(sPath, vbDirectory) <> "" Then FileExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Function FolderExists
' Date  : 10/10/2012
' Type  : Boolean
' Desc  : Checks if a folder exists
' Ex    : FolderExists "C:\Program Files\"
'---------------------------------------------------------------------------------------
Function FolderExists(ByVal sPath As String) As Boolean
    'Add trailing backslash
    If InStr(Len(sPath), sPath, "\") = 0 Then sPath = sPath & "\"
    'If the folder exists return true
    If Dir(sPath, vbDirectory) <> "" Then FolderExists = True
End Function

'---------------------------------------------------------------------------------------
' Proc  : Sub RecMkDir
' Date  : 10/10/2012
' Desc  : Creates an entire directory tree
' Ex    : RecMkDir "C:\Dir1\Dir2\Dir3\"
'---------------------------------------------------------------------------------------
Sub RecMkDir(ByVal sPath As String)
    Dim sDirArray() As String   'Folder names
    Dim sDrive As String        'Base drive
    Dim sNewPath As String      'Path builder
    Dim i As Long               'Counter

    'Add trailing slash
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    'Split at each \
    sDirArray = Split(sPath, "\")
    sDrive = sDirArray(0) & "\"

    'Loop through each directory
    For i = 1 To UBound(sDirArray) - 1
        If Len(sNewPath) = 0 Then
            sNewPath = sDrive & sNewPath & sDirArray(i) & "\"
        Else
            sNewPath = sNewPath & sDirArray(i) & "\"
        End If

        If Not FolderExists(sNewPath) Then
            MkDir sNewPath
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Sub Email
' Date  : 10/11/2012
' Desc  : Sends an email
' Ex    : Email SendTo:=email@example.com, Subject:="example email", Body:="Email Body"
'---------------------------------------------------------------------------------------
Sub Email(SendTo As String, Optional CC As String, Optional BCC As String, Optional Subject As String, Optional Body As String, Optional Attachment As Variant)
    Dim s As Variant              'Attachment string if array is passed
    Dim Mail_Object As Variant    'Outlook application object
    Dim Mail_Single As Variant    'Email object

    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)

    With Mail_Single
        'Add attachments
        Select Case TypeName(Attachment)
            Case "Variant()"
                For Each s In Attachment
                    If s <> Empty Then
                        If FileExists(s) = True Then
                            .attachments.Add s
                        End If
                    End If
                Next
            Case "String"
                If Attachment <> Empty Then
                    If FileExists(Attachment) = True Then
                        .attachments.Add Attachment
                    End If
                End If
        End Select

        'Setup email
        .Subject = Subject
        .To = SendTo
        .CC = CC
        .BCC = BCC
        .HTMLbody = Body
        .Send
    End With
    'Give the email time to send
    Sleep 1500
End Sub

'---------------------------------------------------------------------------------------
' Proc  : Function ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro. Returns true upon success.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps()
    Dim sPath As String     'Gaps file path
    Dim sName As String     'Gaps Sheet Name
    Dim iCounter As Long    'Counter to decrement the date
    Dim iRows As Long       'Total number of rows
    Dim dt As Date          'Date for gaps file name and path
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found
    Dim Gaps As Worksheet           'The sheet named gaps if it exists, else this = nothing
    Dim StartTime As Double         'The time this function was started
    Dim FileFound As Boolean        'Indicates whether or not gaps was found

    StartTime = Timer
    dt = Date - iCounter
    sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
    sName = "3615 " & Format(dt, "m-dd-yy") & ".xlsx"
    FileFound = False

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    Set Gaps = ThisWorkbook.Sheets("Gaps")
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Find gaps
    For iCounter = 0 To 15
        dt = Date - iCounter
        sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        sName = "3615 " & Format(dt, "m-dd-yy") & ".xlsx"
        If FileExists(sPath & sName) Then
            FileFound = True
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If FileFound = True Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            If ThisWorkbook.Sheets("Gaps").Range("A1").Value <> "" Then
                Gaps.Cells.Delete
            End If

            Workbooks.Open sPath & sName
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Gaps").Range("A1")
            ActiveWorkbook.Close

            Sheets("Gaps").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            Columns(1).EntireColumn.Insert
            Range("A1").Value = "SIM"
            Range("A2").Formula = "=C2&D2"
            Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
            Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value

            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Complete"
        Else
            FillInfo FunctionName:="Gaps", _
                     FileDate:=Format(dt, "mm/dd/yy"), _
                     Parameters:="", _
                     ExecutionTime:=Timer - StartTime, _
                     Result:="Failed - User Aborted"
            ERR.Raise 18
        End If
    Else
        MsgBox Prompt:="Gaps could not be found.", Title:="Gaps not found"
        FillInfo FunctionName:="Gaps", _
                 FileDate:=Format(dt, "mm/dd/yy"), _
                 Parameters:="", _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Failed - Gaps not found"
        ERR.Raise 53
    End If

    Application.DisplayAlerts = True
    Exit Sub

CREATE_GAPS:
    ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Gaps"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : FilterSheet
' Date : 1/29/2013
' Desc : Remove all rows that do not match a specified string
'---------------------------------------------------------------------------------------
Sub FilterSheet(sFilter As String, ColNum As Integer, Match As Boolean)
    Dim Rng As Range
    Dim aRng() As Variant
    Dim aHeaders As Variant
    Dim StartTime As Double
    Dim iCounter As Long
    Dim i As Long
    Dim y As Long

    StartTime = Timer
    Set Rng = ActiveSheet.UsedRange
    aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    iCounter = 1

    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

    iCounter = 1
    i = 0
    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ActiveSheet.Cells.Delete
    Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders
    FillInfo "FilterSheet", _
             "", _
             "Filter: " & sFilter & vbCrLf & "Col: " & Columns(ColNum).Address(False, False) & vbCrLf & "Match: " & Match, _
             Timer - StartTime, _
             "Complete"
End Sub


'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range)
    Dim StartTime As Double         'The time this function was started
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    StartTime = Timer
    File = Application.GetOpenFilename()

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File

        ActiveSheet.UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        FillInfo FunctionName:="UserImportFile", _
                 Parameters:="FileName: " & File, _
                 FileDate:=FileDate, _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Complete"

        FillInfo FunctionName:="", _
                 Parameters:="DestRange: " & DestRange.Address(False, False), _
                 Result:="Complete"
    Else
        MsgBox "User aborted file import.", vbOKOnly, "Macro Stopped"
        FillInfo FunctionName:="UserImportFile", _
                 Parameters:="DestRange: " & DestRange.Address(False, False), _
                 ExecutionTime:=Timer - StartTime, _
                 Result:="Failed - User Aborted"
        ThisWorkbook.Activate
        Sheets("Info").Select
        ERR.Raise 18
    End If

End Sub

'---------------------------------------------------------------------------------------
' Proc : FillInfo
' Date : 1/29/2013
' Desc :
'---------------------------------------------------------------------------------------
Sub FillInfo(FunctionName As String, Result As String, Optional ExecutionTime As String = "", Optional Parameters As String = "", Optional FileDate As String = "")
    Dim Info As Worksheet           'Info worksheet if it exists, else this = nothing
    Dim LastSheet As Worksheet      'The previously selected worksheet
    Dim LastWorkbook As Workbook    'The previously activated workbook
    Set LastSheet = ActiveSheet
    Set LastWorkbook = ActiveWorkbook
    Dim Row As Long

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_INFO
    Set Info = ThisWorkbook.Sheets("Info")
    On Error GoTo 0

    ThisWorkbook.Activate
    Sheets("Info").Select
    Range("A1").Value = "Function"
    Range("B1").Value = "Created"
    Range("C1").Value = "Params"
    Range("D1").Value = "Exec Time"
    Range("E1").Value = "Result"

    Row = ActiveSheet.UsedRange.Rows.Count + 1
    Cells(Row, 1).Value = FunctionName
    Cells(Row, 2).Value = FileDate
    Cells(Row, 3).Value = Parameters
    Cells(Row, 4).Value = ExecutionTime
    Cells(Row, 5).Value = Result

    ActiveSheet.UsedRange.Columns.EntireColumn.AutoFit

    LastWorkbook.Activate
    LastSheet.Select
    Exit Sub

CREATE_INFO:
    Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Info"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    
    AddReferences
    codeFolder = CombinePaths(GetWorkbookPath, "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5))
    
    On Error Resume Next
    RecMkDir codeFolder
    On Error GoTo 0

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = CombinePaths(codeFolder, comp.Name & ".bas")
                DeleteFile FileName
                comp.Export FileName
            Case 2
                FileName = CombinePaths(codeFolder, comp.Name & ".cls")
                DeleteFile FileName
                comp.Export FileName
            Case 3
                FileName = CombinePaths(codeFolder, comp.Name & ".frm")
                DeleteFile FileName
                comp.Export FileName
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFile
' Date : 3/19/2013
' Desc : Deletes a file
'---------------------------------------------------------------------------------------
Sub DeleteFile(FileName As String)
    On Error Resume Next
    Kill FileName
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : CombinePaths
' Date : 3/19/2013
' Desc : Adds folders onto the end of a file path
'---------------------------------------------------------------------------------------
Function CombinePaths(ByVal Path1 As String, ByVal Path2 As String) As String
    If Not EndsWith(Path1, "\") Then
        Path1 = Path1 & "\"
    End If
    CombinePaths = Path1 & Path2
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds references required for helper functions
'---------------------------------------------------------------------------------------
Sub AddReferences()
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean

    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = "{0002E157-0000-0000-C000-000000000046}" And Ref.Major = 5 And Ref.Minor = 3 Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid "{0002E157-0000-0000-C000-000000000046}", 5, 3
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes references required for helper functions
'---------------------------------------------------------------------------------------
Sub RemoveReferences()
    Dim Ref As Variant

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = "{0002E157-0000-0000-C000-000000000046}" And Ref.Major = 5 And Ref.Minor = 3 Then
            Application.VBE.ActiveVBProject.References.Remove Ref
        End If
    Next
End Sub

