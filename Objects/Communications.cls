VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
GroupMember.Show
End Sub

Private Sub CommandButton2_Click()
Dim rng As Range
    
    On Error Resume Next
    With Selection.Cells(1)
        Set rng = Intersect(.EntireRow, ActiveCell.ListObject.DataBodyRange)
        On Error GoTo 0
        If rng Is Nothing Then
            MsgBox "Please select a valid table cell.", vbCritical
        Else
            rng.Delete xlShiftUp
        End If
    End With

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
