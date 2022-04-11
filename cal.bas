Attribute VB_Name = "Module1"
Option Explicit

Sub cal()
Range("E1").Value = Range("A1").Value + Range("C1").Value
Range("E2").Value = Range("A1").Value - Range("C1").Value
Range("E3").Value = Range("A1").Value * Range("C1").Value
Range("E4").Value = Range("A1").Value / Range("C1").Value
End Sub

Sub calNew()

Dim v1, v2 As Integer
v1 = Range("A1").Value
v2 = Range("C1").Value
Range("E1").Value = v1 + v2
Range("E2").Value = v1 - v2
Range("E3").Value = v1 * v2
Range("E4").Value = v1 / v2
End Sub
