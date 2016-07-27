VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectPeriods 
   Caption         =   "Select Class Periods"
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   -7365
   ClientWidth     =   7005
   OleObjectBlob   =   "SelectPeriods.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectPeriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DidSelect As Boolean

Private Sub CheckBox1_Click()

Dim counter As Variant

If SelectPeriods.CheckBox1.Value = True Then

    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
    
        SelectPeriods.ListBox1.Selected(counter) = True
        
    Next counter
    
    SelectPeriods.CommandButton2.Enabled = True
    
Else

    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
    
        SelectPeriods.ListBox1.Selected(counter) = False
        
    Next counter
    
    SelectPeriods.CommandButton2.Enabled = False

End If

Call EnableButton

End Sub

Private Sub CommandButton1_Click()

DidSelect = False
SelectPeriods.ListBox1.Clear
SelectPeriods.Hide

End Sub

Private Sub CommandButton2_Click()

DidSelect = True
SelectPeriods.ListBox1.Clear
SelectPeriods.Hide

End Sub

Private Sub Label1_Click()

End Sub

Private Sub ListBox1_Change()

Dim counter As Long
Dim EmptyOrNot As New Collection

    For counter = 0 To SelectPeriods.ListBox1.ListCount - 1
    
        If SelectPeriods.ListBox1.Selected(counter) = True Then
        
            EmptyOrNot.Add (SelectPeriods.ListBox1.List(counter))
            
        End If
        
    Next counter
        
    If EmptyOrNot.Count = 0 Then
    
        SelectPeriods.CommandButton2.Enabled = False
        
    Else:
    
        SelectPeriods.CommandButton2.Enabled = True
        
    End If
    
    Call EnableButton

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub EnableButton()

    If SelectPeriods.CommandButton2.Enabled = False Then
    
        SelectPeriods.CommandButton2.BackColor = vbButtonShadow
        
    Else:
    
        SelectPeriods.CommandButton2.BackColor = vbHighlight
        
    End If

End Sub
