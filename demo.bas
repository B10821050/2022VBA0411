Attribute VB_Name = "Module1"

Sub FirstVB()
'�i�ߤH�A0411

MsgBox ("�ڪ��Ĥ@��ۦ�}�o��VBA ")

End Sub

Sub aa()

Range("A1:C9").Value = "vba"
Range("E1").Value = "�ڷR�g�{��"

End Sub

Sub democells()

cells(1, 5).Value = "�ڷR�g�{��"
cells(1, "E").Value = "�ڷR�g�{��"
End Sub

Sub democells2()

cells(1, 7).Value = "�ڷR�g�{��"
cells(2, "G").Value = "�ڷR�g�{��"
Range("G3").Value = "�ڷR�g�{��"
End Sub

Sub demoTime()

cells(1, 6).Value = Now()
cells(1, "F").Value = Now()
End Sub

Sub demoTimeclear()

cells(1, 6).Clear
cells(1, "F").Clear

End Sub

Sub Stringdemo()

Dim i As String
i = "�i�D�����ڷR�gVBA"
cells(1, 8).Value = i


End Sub

Sub intdemo()

Dim j As Integer
j = 1000
cells(2, "H").Value = j

End Sub

Sub Singledemo()

Dim k As Single
k = 878.696
cells(3, 9).Value = k

End Sub

Sub Doubledemo()

Dim m As Double
m = 787.145387597238
cells(4, 8).Value = m

End Sub

Sub Datedemo()

Dim n As Date
n = Now
cells(5, 8).Value = n

End Sub

Sub Booleandemo()

Dim o As Boolean
o = True
cells(6, 8).Value = o

End Sub

