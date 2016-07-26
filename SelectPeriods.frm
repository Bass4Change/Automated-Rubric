VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectPeriods 
   Caption         =   "Select Class Periods"
   ClientHeight    =   8000
   ClientLeft      =   0
   ClientTop       =   -7360
   ClientWidth     =   7000
   OleObjectBlob   =   "SelectPeriods.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectPeriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

Dim counter As Variant

If SelectPeriods.CheckBox1.Value = True Then

    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
    
        SelectPeriods.ListBox1.Selected(counter) = True
        
    Next counter
    
Else

    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
    
        SelectPeriods.ListBox1.Selected(counter) = False
        
    Next counter

End If

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
Private Sub CommandButton1_Click()

DidSelect = False
Unload Me
SelectPeriods.Hide

End Sub

Private Sub CommandButton2_Click()

DidSelect = True
SelectPeriods.Hide

End Sub

Private Sub Label1_Click()

End Sub

Private Sub ListBox1_Click()


End Sub

Private Sub UserForm_Click()

End Sub
