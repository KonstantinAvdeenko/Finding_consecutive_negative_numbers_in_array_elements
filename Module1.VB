Private Sub CommandButton1_Click()
'каждому элементу массива присваивается рандомное положительное или отрицательное число'
For i = 1 To 30
Cells(1, i) = Int((100 * Rnd) - 50)
Next i
End Sub

Private Sub CommandButton2_Click()
'подсчет максимального количества подряд идущих отрицательных элементов массива'
For i = 1 To 30
If Cells(1, i) < 0 Then
k = k + 1
Else
k = 0
End If
If n < k Then
n = k
End If
Next i
MsgBox (n)
End Sub

Private Sub CommandButton3_Click()
'закрытие формы'
UserForm1.Hide
End Sub