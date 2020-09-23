Attribute VB_Name = "Engine"
Public Ver(5000, 2) As Single
Public TempV(5000, 2) As Single
Public Lin(5000, 5) As Single
Public Sa As Double, TransX As Double, TransY As Double
Public TransZ As Double, RotX As Double, RotY As Double, RotZ As Double
Public Zm As Double, VerNum As Double, LineNum As Double
Public ViewDis As Double
Public ZCenter(5000) As Double
Public XOrigin As Double, YOrigin As Double

Public Sub Change()
Convert
Sca (Sa)
Translate (TransX), (TransY), (TransZ)
Rotate (RotX), (RotY), (RotZ)
Zoom (Zm)
Transform
SortZ
ImageView
End Sub
Public Sub Convert()
Dim a, b
For a = 0 To VerNum
For b = 0 To 2
TempV(a, b) = Ver(a, b)
Next b, a
End Sub
Public Sub Sca(S As Double)
Dim a, b
For a = 0 To VerNum
For b = 0 To 2
TempV(a, b) = TempV(a, b) * S
Next b, a
End Sub
Public Sub Translate(X As Double, Y As Double, Z As Double)
Dim a
For a = 0 To VerNum
TempV(a, 0) = TempV(a, 0) + X
TempV(a, 1) = TempV(a, 1) + Y
TempV(a, 2) = TempV(a, 2) + Z
Next a
End Sub
Public Sub Rotate(X As Double, Y As Double, Z As Double)
Dim Xn As Double, Yn As Double, Zn As Double
Xn = X * (3.141592654 / 180)
Yn = Y * (3.141592654 / 180)
Zn = Z * (3.141592654 / 180)
Dim a, X1 As Double, Y1 As Double, Z1 As Double
For a = 0 To VerNum
X = TempV(a, 0): Y = TempV(a, 1): Z = TempV(a, 2)
X1 = X * Cos(Zn) - Y * Sin(Zn)
Y1 = Y * Cos(Zn) + X * Sin(Zn)
Z1 = Z
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

X = TempV(a, 0): Y = TempV(a, 1): Z = TempV(a, 2)
Y1 = Y * Cos(Xn) - Z * Sin(Xn)
Z1 = Y * Sin(Xn) + Z * Cos(Xn)
X1 = X
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

X = TempV(a, 0): Y = TempV(a, 1): Z = TempV(a, 2)
Z1 = Z * Cos(Yn) - X * Sin(Yn)
X1 = Z * Sin(Yn) + X * Cos(Yn)
Y1 = Y
TempV(a, 0) = X1: TempV(a, 1) = Y1: TempV(a, 2) = Z1

Next a
End Sub
Public Sub Zoom(Z As Double)
Dim a
For a = 0 To VerNum
TempV(a, 2) = TempV(a, 2) + Z
Next a
End Sub
Public Sub Transform()
Dim X As Double, Y As Double, Z As Double
Dim a
For a = 0 To VerNum
X = TempV(a, 0): Y = TempV(a, 1): Z = TempV(a, 2)
If Z < -999 Then Z = -999
If Z > 999 Then Z = 999
Z = Z + 1000
Z = 2000 - Z
X = (X / Z) * 1000
Y = (Y / Z) * 1000
TempV(a, 0) = XOrigin + X
TempV(a, 1) = YOrigin + Y
Next a
End Sub
Public Sub SortZ()
Dim Z1 As Double, Z2 As Double, ZMin As Double, ZMax As Double
Dim a, Dummy As Double, Sw
For a = 0 To LineNum
Z1 = TempV(Lin(a, 0), 2)
Z2 = TempV(Lin(a, 1), 2)
If Z1 < Z2 Then ZMin = Z1: ZMax = Z2
If Z1 > Z2 Then ZMin = Z2: ZMax = Z1
If Z1 = Z2 Then ZMin = Z1: ZMax = Z2
ZCenter(a) = (ZMax - ZMin) / 2
ZCenter(a) = ZMin + ZCenter(a)
Next a
Dim b
ZStart:
Sw = 0
For a = 0 To LineNum - 1
If ZCenter(a) > ZCenter(a + 1) Then
For b = 0 To 4
Dummy = Lin(a, b): Lin(a, b) = Lin(a + 1, b): Lin(a + 1, b) = Dummy
Next b
Dummy = ZCenter(a): ZCenter(a) = ZCenter(a + 1): ZCenter(a + 1) = Dummy
Sw = 1
End If
Next a
If Sw <> 0 Then GoTo ZStart
End Sub
Public Sub ImageView()
Dim Aa, Bb
Dim X1 As Double, Y1 As Double, Z1 As Double
Dim X2 As Double, Y2 As Double, Z2 As Double
Dim C1, C2, C3
Dim a
View.Display.Cls
For a = 0 To LineNum
Aa = Lin(a, 0): Bb = Lin(a, 1)
X1 = TempV(Aa, 0): Y1 = TempV(Aa, 1): Z1 = TempV(Aa, 2)
X2 = TempV(Bb, 0): Y2 = TempV(Bb, 1): Z2 = TempV(Bb, 2)
If Z1 > ViewDis And Z2 > ViewDis Then GoTo Skip
C1 = Lin(a, 2): C2 = Lin(a, 3): C3 = Lin(a, 4)
View.Display.Line (X1, Y1)-(X2, Y2), RGB(C1, C2, C3)
Skip:
Next a



End Sub
