Attribute VB_Name = "MasterCheckbox"

Sub BetaAutomatedRubric_MasterCheckbox_Click()

    Dim CB As CheckBox

    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("Master Checkbox").Name Then
            CB.Value = ActiveSheet.CheckBoxes("Master Checkbox").Value
        End If
    Next CB
    
End Sub

Sub MixedState()

    Dim CB As CheckBox
    For Each CB In ActiveSheet.CheckBoxes
        If CB.Name <> ActiveSheet.CheckBoxes("Master Checkbox").Name And CB.Value <> ActiveSheet.CheckBoxes("Master Checkbox").Value And ActiveSheet.CheckBoxes("Master Checkbox").Value <> 2 Then
            ActiveSheet.CheckBoxes("Master Checkbox").Value = 2
        Exit For
            Else
            ActiveSheet.CheckBoxes("Master Checkbox").Value = CB.Value
        End If
    Next CB
End Sub

Sub AssignMacro()

    Dim CB As CheckBox
        For Each CB In ActiveSheet.CheckBoxes
                CB.OnAction = ""
            
        Next CB
                
End Sub
Sub AssignOneMacro()
Attribute AssignOneMacro.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AssignOneMacro Macro
'

'
    ActiveSheet.Shapes.Range(Array("Check Box 305")).Select
    Selection.OnAction = "MixedState"
End Sub
