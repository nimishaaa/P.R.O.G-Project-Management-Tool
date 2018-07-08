VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

Dim ws As Worksheet

Dim i As Integer

Dim o As Integer

Set ws = ThisWorkbook.Sheets("Task Manager")

i = Cells(Rows.Count, "M").End(xlUp).Row

For i = 1 To 20

o = ws.Cells(Rows.Count, "O").End(xlUp).Row
    If o = 2 And Cells(o, "O").Value = "" Then
    o = o
    Else
    o = o + 1
    End If

 If Sheet4.Cells(i, "M").Value = "Completed" Then
 Sheet4.Cells(o, "O").Value = Sheet4.Cells(i, "J").Value
 Sheet4.Cells(o, "P").Value = Sheet4.Cells(i, "L").Value
 
 Sheet4.Cells(i, "J").Value = ""
 Sheet4.Cells(i, "K").Value = ""
 Sheet4.Cells(i, "L").Value = ""
 Sheet4.Cells(i, "M").Value = ""
 
Sheet4.Cells(i, "J").Delete Shift:=xlUp

 End If
 
 Next i
 
 
End Sub

Private Sub CommandButton2_Click()
Dim ws As Worksheet

Dim i As Integer

Dim j As Integer


Set ws = ThisWorkbook.Sheets("Task Manager")

For i = 1 To 20

j = ws.Cells(Rows.Count, "J").End(xlUp).Row
    If j = 2 And Cells(j, "J").Value = "" Then
    j = j
    Else
    j = j + 1
    End If

 If Sheet4.Cells(i, "G").Value = "On Going" Then
 Sheet4.Cells(j, "J").Value = Sheet4.Cells(i, "B").Value
 Sheet4.Cells(j, "K").Value = Sheet4.Cells(i, "E").Value
 Sheet4.Cells(j, "L").Value = Sheet4.Cells(i, "F").Value
 Sheet4.Cells(j, "M").Value = Sheet4.Cells(i, "G").Value
 
 
 Sheet4.Cells(i, "D").Value = ""
 Sheet4.Cells(i, "B").Value = ""
 Sheet4.Cells(i, "C").Value = ""
 Sheet4.Cells(i, "E").Value = ""
 Sheet4.Cells(i, "F").Value = ""
 Sheet4.Cells(i, "G").Value = ""
 
 Sheet4.Cells(i, "B").Delete Shift:=xlUp
 
 End If
 
 Next i
 
End Sub

Private Sub CommandButton3_Click()
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

Private Sub CommandButton4_Click()
TaskManager.Show
End Sub

