Attribute VB_Name = "CreateRubricsCode"
Public DidSelect As Boolean

Public iRow As Long
Public counter As Long

Public CellName As String

Public NewBook As Workbook

Public ClassRoster As Worksheet
Public StartHere As Worksheet

Public ClassPeriods As New Collection
Public TheBooks As New Collection

Sub CreateRubrics()
Attribute CreateRubrics.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Silly As Boolean

'I would like to resolve this Screen Updating issue

Application.ScreenUpdating = False

Set StartHere = Workbooks("Automated Rubric.xlsm").Worksheets("Start Here")
StartHere.Activate

Silly = True
DidSelect = True

If StartHere.Range("A2") = "Step 1: Select this cell (A2)." Then
    
    Silly = False
    
Else:

    Silly = True
    
End If
    
If Silly = True Then

    'Subroutine creates a new workbook to format all relevant class rosters.
    'The rest of the procedures will reference this class roster, but by the end of the program,
    'this workbook will be deleted and not saved.
    
    Call FormatRoster
    Call AddDefinitiveNumbers
    
    'UserForm to save rubric workbooks to desired folder (for macs)
    
    SelectPeriods.Show

    If DidSelect = False Then
    
        Call CloseRoster
        Exit Sub
        
    End If
    
    'File Saving Stage
    
    'Next sub is to reduce roster to relevant students
    
    Call ReduceStudents
    
    'Next sub retrieves file location and name from user
    
    Call TheFileName
             
    'Next sub checks to make sure all student names are compatible as Worksheet names
    
    Call NameCheck
    
    'Looks like I accidently deleted the remove period formula.
    'I'll need to replace that.
    
    'Next sub copies and pastes rubrics into the new workbooks.
    
    Call CopyPasta
    
    Call CloseRoster
    
    Call SaveTheBooks
        
Else:

    MsgBox ("You didn't paste your Weekly Attendance Roster, silly! Try again.")
    
End If
    
End Sub

Private Sub FormatRoster()
    
    Dim OldBook As Workbook
    Dim NewSheet As Worksheet
    
    Dim period As Range
    Dim NewPeriod As Range
    Dim myRange As Range
    Dim FirstName As Range
    
    Dim myInt As Integer
    
    For Each OldBook In Workbooks
    
        If OldBook.Name <> "Automated Rubric.xlsm" Then
            OldBook.Saved = True
            OldBook.Close
            
        End If
        
    Next OldBook
    
    Set NewBook = Workbooks.Add
                            
                    StartHere.Copy Before:=NewBook.Worksheets("Sheet1")
                    
                    With NewBook
                        
                        Application.DisplayAlerts = False
                        Worksheets("Sheet1").Delete
                        Application.DisplayAlerts = True
                        Set NewSheet = NewBook.ActiveSheet
                        NewSheet.Name = "Class Roster"
                        
                    End With
    
    iRow = 2
        
    Columns("A:A").Delete Shift:=xlToLeft
    Columns("D:AD").Delete Shift:=xlToLeft
    
    Rows(iRow & ":" & iRow).Delete Shift:=xlUp
    
    Cells(iRow, 2) = "Yes"
        
    Do Until IsEmpty(Cells(iRow, 2)) And IsEmpty(Cells(iRow, 1))
    
        Cells(iRow, 2) = "Yes"
        Rows(iRow & ":" & iRow).Delete Shift:=xlUp
        Set period = Cells(iRow, 1)
        Set NewPeriod = Cells(iRow, 2)
        period.Select
        period.Cut Destination:=Range("B" & iRow)
        Set NewPeriod = Cells(iRow, 2)
        iRow = iRow + 2
        Set FirstName = Cells(iRow, 2)
        FirstName.Select
        
        Do Until IsEmpty(Cells(iRow, 2))
        
            iRow = iRow + 1
            Cells(iRow, 2).Select
        
        Loop
        
        Set myRange = Range("A:A")
        myInt = WorksheetFunction.CountA(myRange)
           
        Do While IsEmpty(Cells(iRow, 1)) And myInt > 0
        
            Rows(iRow & ":" & iRow).Delete Shift:=xlUo
            
        Loop
        
    Loop
    
    Columns("A:A").Delete Shift:=xlToLeft
    iRow = 2
    Cells(iRow, 1).Select
    
    Do Until IsEmpty(Cells(iRow, 1))
    
        If Not Cells(iRow, 1) = "Student Name" Then
            
            iRow = iRow + 1
            Cells(iRow, 1).Select
            
        Else:
        
            Rows(iRow & ":" & iRow).Delete Shift:=xlUp
            Cells(iRow, 1).Select
    
        End If
    
    Loop
    
    iRow = 2
    
    Do Until IsEmpty(Cells(iRow, 1))
    
        If Not Cells(iRow, 2) = "Strategic Supp" Then
        
            iRow = iRow + 1
            Cells(iRow, 2).Select
            
        Else:
        
            Do Until IsEmpty(Cells(iRow, 1))
            
                Rows(iRow & ":" & iRow).Delete Shift:=xlUp
                Cells(iRow, 2).Select
                
            Loop
            
        End If
        
    Loop
    
    iRow = 2
    Set ClassRoster = NewBook.Worksheets("Class Roster")
    
End Sub

Private Sub AddDefinitiveNumbers()
    
    Dim DefinitiveNumber As Long
    Dim ValidClassSection As String
    Dim CourseName As String
    
    'Add list of definitive numbers to the left
    ClassRoster.Columns("A:A").Insert Shift:=xlTotheLeft
    Cells(iRow, 2).Activate
    
    DefinitiveNumber = 1

    Do Until IsEmpty(ClassRoster.Cells(iRow, 2))
        
            If ClassRoster.Cells(iRow, 3) = "" Then
                
                iRow = iRow + 1
            
            Else:
                
                ClassRoster.Cells(iRow, 1) = DefinitiveNumber
                ValidClassSection = ClassRoster.Cells(iRow, 2)
                CourseName = ClassRoster.Cells(iRow, 3)
                SelectPeriods.ListBox1.AddItem ("Period " + ValidClassSection + " - " + CourseName)
                
                DefinitiveNumber = DefinitiveNumber + 1
                iRow = iRow + 1
                
            End If
            
    Loop

End Sub

Private Sub ReduceStudents()
    
    Dim ValidPeriod As String
    
    iRow = 2
    
    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
        
        If SelectPeriods.ListBox1.Selected(counter) = True Then
                
                ValidPeriod = SelectPeriods.ListBox1.List(counter)
                ClassPeriods.Add (ValidPeriod)
            
            Do Until ClassRoster.Cells(iRow, 1) <> counter + 1 And IsEmpty(ClassRoster.Cells(iRow, 1)) = False Or IsEmpty(ClassRoster.Cells(iRow, 2)) = True
            
                    iRow = iRow + 1
                    
            Loop
                
            'Condition following is to eliminate students no longer needed
        
        Else

                Do Until ClassRoster.Cells(iRow, 1) > counter + 1 Or IsEmpty(ClassRoster.Cells(iRow, 2))
                    
                    ClassRoster.Rows(iRow & ":" & iRow).Delete xlUp
                        
                Loop
                    
        End If
            
    Next counter
        
    ClassRoster.Columns("A:A").Delete xlLeft

End Sub



Private Sub NameCheck()

    Dim RosterName As Range
    Dim Burrito As Boolean
    
    iRow = 3
    
    Do Until IsEmpty(ClassRoster.Cells(iRow, 1))
    
        Set RosterName = ClassRoster.Cells(iRow, 1)
        CellName = RosterName
        
        If IsGoodAscii(CellName) = False Then
        
            Burrito = False
        
            Do While Burrito = False
            
            CellName = InputBox("It looks like the name " & CellName & "has an invalid character. How would you like to respell the name?")
            
                If CellName = vbNullString Then
                    
                    CellName = ClassRoster.Cells(iRow, 1)
                    Burrito = False
                    
                Else:
                
                    If IsGoodAscii(CellName) = True Then
                        
                        Burrito = True
                        
                    Else:
                    
                        Burrito = False
                    
                    End If
                    
                End If
            
            Loop
            
        Else:
        
            CellName = CellName
            
        End If
        
        iRow = iRow + 1
        
    Loop
    
End Sub

Private Sub TheFileName()
        
        Dim potato As String
        Dim wkb1 As Workbook
        
        For counter = 1 To ClassPeriods.Count
    
            potato = ClassPeriods(counter)
            Workbooks.Add (1)
            MacGetSaveAsFilenameExcel MyInitialFilename:=potato, FileExtension:="xlsm"
            Set wkb1 = ActiveWorkbook
            TheBooks.Add Item:=wkb1
        
        Next counter
    
End Sub

Private Sub CopyPasta()

    Dim ReName As Range
    Dim CellNumber As Boolean
    
    Dim wkb As Workbook
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Set ws1 = Workbooks("Automated Rubric.xlsm").Worksheets("Beta Automated Rubric")
            
    iRow = 3
    
    For Each wkb In TheBooks
        
        CellNumber = True
        Do While CellNumber = True
                            
                    CellName = ClassRoster.Cells(iRow, 1)
                    ws1.Copy Before:=wkb.Worksheets("Sheet1")
                    Set ws2 = wkb.ActiveSheet
                    ws2.Name = CellName
                    Set ReName = ws2.Cells(2, 1)
                    ReName = "Name: " & CellName
                    ReName.Font.Size = 12
                    
                    If IsEmpty(ClassRoster.Cells(iRow + 1, 2)) And Not IsEmpty(ClassRoster.Cells(iRow + 1, 1)) Then
                
                        CellNumber = True
                        iRow = iRow + 1
                            
                    Else:
                        
                        CellNumber = False
                        iRow = iRow + 2
                    
                    End If
                    
        Loop
            
    Next wkb
    
End Sub

Private Sub SaveTheBooks()

Dim DifBook As Workbook
Dim EndAll As Worksheet

    For Each DifBook In Workbooks
    
        For Each EndAll In DifBook.Worksheets
        
            If EndAll.Name = "Start Here" Or EndAll.Name = "Beta Automated Rubric" Then
            
                Exit For
                
            Else:
            
                If EndAll.Name = "Sheet1" Then
            
                    Application.DisplayAlerts = False
                    EndAll.Delete
                    Application.DisplayAlerts = True
                    
                End If
            
            End If
                    
        Next EndAll
        
    Next DifBook

    For Each DifBook In Workbooks
    
        DifBook.Save
        
    Next DifBook

End Sub
Private Sub CloseRoster()

    NewBook.Saved = True
    Application.ScreenUpdating = True
    NewBook.Close

End Sub

Function MacGetSaveAsFilenameExcel(MyInitialFilename As String, FileExtension As String)
'Ron de Bruin, 03-April-2015
'Custom function for the Mac to save the activeworkbook in the format you want.
'If FileExtension = "" you can save in the following formats : xls, xlsx, xlsm, xlsb
'You can also set FileExtension to the extension you want like "xlsx" for example
    Dim FName As Variant
    Dim FileFormatValue As Long
    Dim TestIfOpen As Workbook
    Dim FileExtGetSaveAsFilename As String

Again:         FName = False
    'Call VBA GetSaveAsFilename
    'Note: InitialFilename is the only parameter that works on a Mac
    FName = Application.GetSaveAsFilename(InitialFileName:=MyInitialFilename)

    If FName <> False Then
        'Get the file extension
        FileExtGetSaveAsFilename = LCase(Right(FName, Len(FName) - InStrRev(FName, ".", , 1)))

        If FileExtension <> "" Then
            If FileExtension <> FileExtGetSaveAsFilename Then
                MsgBox "You didn't follow the instructions! Please save the file in this format : " & FileExtension
                GoTo Again
            End If
            If ActiveWorkbook.HasVBProject = True And LCase(FileExtension) = "xlsx" Then
                MsgBox "Your workbook has VBA code, please not save in xlsx format"
                Exit Function
            End If
        Else
            If ActiveWorkbook.HasVBProject = True And LCase(FileExtGetSaveAsFilename) = "xlsx" Then
                MsgBox "Your workbook has VBA code, please not save in xlsx format"
                GoTo Again
            End If
        End If

        'Find the correct FileFormat that match the choice in the "Save as type" list
        'and set the FileFormatValue, Extension and FileFormatValue must match.
        'Note : You can add or delete items to/from the list below if you want.
        Select Case FileExtGetSaveAsFilename
        Case "xls": FileFormatValue = 57
        Case "xlsx": FileFormatValue = 52
        Case "xlsm": FileFormatValue = 53
        Case "xlsb": FileFormatValue = 51
        Case Else: FileFormatValue = 0
        End Select
        If FileFormatValue = 0 Then
            MsgBox "Sorry, FileFormat not allowed"
            GoTo Again
        Else
            'Error check if there is a file open with that name
            Set TestIfOpen = Nothing
            On Error Resume Next
            Set TestIfOpen = Workbooks(LCase(Right(FName, Len(FName) - InStrRev(FName, _
                Application.PathSeparator, , 1))))
            On Error GoTo 0

            If Not TestIfOpen Is Nothing Then
                MsgBox "You are not allowed to overwrite a file that is open with the same name, " & _
                "use a different name or close the file with the same name first."
                GoTo Again
            End If
        End If

        'Now we have the information to Save the file
        Application.DisplayAlerts = False
        On Error Resume Next
        ActiveWorkbook.SaveAs FName, FileFormat:=FileFormatValue
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If

End Function

Function IsGoodAscii(aString As String) As Boolean
Dim i As Long
Dim iLim As Long
i = 1
iLim = Len(aString)

While i <= iLim
    If (Asc(Mid(aString, i, 1)) < 48 And Asc(Mid(aString, i, 1)) > 32 And Not Asc(Mid(aString, i, 1)) = 46 And Not Asc(Mid(aString, i, 1)) = 44 And Not Asc(Mid(aString, i, 1)) = 45) Xor (Asc(Mid(aString, i, 1)) > 90 And Asc(Mid(aString, i, 1)) < 96) Xor Asc(Mid(aString, i, 1)) > 122 Then
        IsGoodAscii = False
        Exit Function
    End If
    i = i + 1
Wend

IsGoodAscii = True
End Function


