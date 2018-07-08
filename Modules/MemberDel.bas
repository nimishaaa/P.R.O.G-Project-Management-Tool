Attribute VB_Name = "Module6"
Sub DeleteMember()
Attribute DeleteMember.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' DeleteMember Macro
'
' Keyboard Shortcut: Ctrl+d
'
    Range("A2:G2").Select
    Selection.Delete Shift:=xlUp
End Sub
