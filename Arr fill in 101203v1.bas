Attribute VB_Name = "Ä£¿é5"
Sub sztest_1()
    Dim arr(1 To 10) As Integer, i, m As Integer, a1, a2, a3 As String
    Let i = 1
    For i = 1 To 10
        Let arr(i) = i
        a1 = "arr("
        a2 = ")="
        a3 = a1 & i & a2 & i
        ActiveCell(1 + i, 1 - 1).Value = a3
        'ActiveCell(1 + i, 1).Value = arr(i)
    Next
End Sub

