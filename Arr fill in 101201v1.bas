Attribute VB_Name = "Ä£¿é3"
Sub sztest_1()
    Dim arr(1 To 10) As Integer, i, m As Integer
    Let i = 1
    For i = 1 To 10
        Let arr(i) = i
    Next
    Range("A1") = arr(1)
    Range("A2") = arr(2)
    Range("A3") = arr(3)
    Range("A4") = arr(4)
    Range("A5") = arr(5)
    Range("A6") = arr(6)
    Range("A7") = arr(7)
    Range("A8") = arr(8)
    Range("A9") = arr(9)
    Range("A10") = arr(10)
End Sub
