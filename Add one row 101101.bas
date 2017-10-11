Attribute VB_Name = "Ä£¿é1"
Sub Add_row()
Attribute Add_row.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Add_row ºê
'

'
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    ActiveCell.Offset(2, 0).Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown
    ActiveCell.Select
End Sub
