VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   1740
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   1740
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************
'*                               *
'* OOO    OOO   OOO  OOO    OOO  *
'* O  O  O   O O   O O  O  O   O *
'* OOO   O   O O   O O  O  OOOOO *
'* O  O  O   O O   O O  O  O   O *
'* OOO    OOO   OOO  OOO   O   O *
'*                               *
'*********************************
'
' 3D Demo
' September 99
' Feel free to use this code, just remember to give credit
' where credit is due.....I would do the same for you.
'
' Da Booda
' Any comments or questions...
' Email boodaone@aol.com
'
' P.S.  Please forgive the roughness of my code, I
' don't consider myself a professional...and if any one
' could better this program, feel free and please EMail me
' the changes.


Private Sub Form_Load()
View.Show
Remote.Show
Sa = 1
TransX = 0: TransY = 0: TransZ = 0
RotX = 0: RotY = 0: RotZ = 0
Zm = 0
VerNum = 25
LineNum = 41
ViewDis = 999
XOrigin = 400
YOrigin = 300

Rem Vertices
Ver(0, 0) = -10: Ver(0, 1) = -10: Ver(0, 2) = -10
Ver(1, 0) = 10: Ver(1, 1) = -10: Ver(1, 2) = -10
Ver(2, 0) = 10: Ver(2, 1) = 10: Ver(2, 2) = -10
Ver(3, 0) = -10: Ver(3, 1) = 10: Ver(3, 2) = -10

Ver(4, 0) = -10: Ver(4, 1) = -10: Ver(4, 2) = 10
Ver(5, 0) = 10: Ver(5, 1) = -10: Ver(5, 2) = 10
Ver(6, 0) = 10: Ver(6, 1) = 10: Ver(6, 2) = 10
Ver(7, 0) = -10: Ver(7, 1) = 10: Ver(7, 2) = 10

Ver(8, 0) = -7: Ver(8, 1) = -7: Ver(8, 2) = -7
Ver(9, 0) = 7: Ver(9, 1) = -7: Ver(9, 2) = -7
Ver(10, 0) = 7: Ver(10, 1) = 7: Ver(10, 2) = -7
Ver(11, 0) = -7: Ver(11, 1) = 7: Ver(11, 2) = -7

Ver(12, 0) = -7: Ver(12, 1) = -7: Ver(12, 2) = 7
Ver(13, 0) = 7: Ver(13, 1) = -7: Ver(13, 2) = 7
Ver(14, 0) = 7: Ver(14, 1) = 7: Ver(14, 2) = 7
Ver(15, 0) = -7: Ver(15, 1) = 7: Ver(15, 2) = 7

Ver(16, 0) = 0: Ver(16, 1) = -10: Ver(16, 2) = 0
Ver(17, 0) = -10: Ver(17, 1) = 0: Ver(17, 2) = 0
Ver(18, 0) = 0: Ver(18, 1) = 0: Ver(18, 2) = -10
Ver(19, 0) = 10: Ver(19, 1) = 0: Ver(19, 2) = 0
Ver(20, 0) = 0: Ver(20, 1) = 0: Ver(20, 2) = 10
Ver(21, 0) = 0: Ver(21, 1) = 10: Ver(21, 2) = 0

Ver(22, 0) = -7: Ver(22, 1) = -7: Ver(22, 2) = 10
Ver(23, 0) = 7: Ver(23, 1) = -7: Ver(23, 2) = 10
Ver(24, 0) = 7: Ver(24, 1) = 7: Ver(24, 2) = 10
Ver(25, 0) = -7: Ver(25, 1) = 7: Ver(25, 2) = 10

Rem Lines
Lin(0, 0) = 0: Lin(0, 1) = 1
Lin(1, 0) = 1: Lin(1, 1) = 2
Lin(2, 0) = 2: Lin(2, 1) = 3
Lin(3, 0) = 3: Lin(3, 1) = 0
Lin(4, 0) = 4: Lin(4, 1) = 5
Lin(5, 0) = 5: Lin(5, 1) = 6
Lin(6, 0) = 6: Lin(6, 1) = 7
Lin(7, 0) = 7: Lin(7, 1) = 4
Lin(8, 0) = 0: Lin(8, 1) = 4
Lin(9, 0) = 1: Lin(9, 1) = 5
Lin(10, 0) = 2: Lin(10, 1) = 6
Lin(11, 0) = 3: Lin(11, 1) = 7

Lin(12, 0) = 8: Lin(12, 1) = 9
Lin(13, 0) = 9: Lin(13, 1) = 10
Lin(14, 0) = 10: Lin(14, 1) = 11
Lin(15, 0) = 11: Lin(15, 1) = 8
Lin(16, 0) = 12: Lin(16, 1) = 13
Lin(17, 0) = 13: Lin(17, 1) = 14
Lin(18, 0) = 14: Lin(18, 1) = 15
Lin(19, 0) = 15: Lin(19, 1) = 12
Lin(20, 0) = 8: Lin(20, 1) = 12
Lin(21, 0) = 9: Lin(21, 1) = 13
Lin(22, 0) = 10: Lin(22, 1) = 14
Lin(23, 0) = 11: Lin(23, 1) = 15

Lin(24, 0) = 16: Lin(24, 1) = 17
Lin(25, 0) = 16: Lin(25, 1) = 18
Lin(26, 0) = 16: Lin(26, 1) = 19
Lin(27, 0) = 16: Lin(27, 1) = 20
Lin(28, 0) = 17: Lin(28, 1) = 18
Lin(29, 0) = 18: Lin(29, 1) = 19
Lin(30, 0) = 19: Lin(30, 1) = 20
Lin(31, 0) = 20: Lin(31, 1) = 17
Lin(32, 0) = 21: Lin(32, 1) = 17
Lin(33, 0) = 21: Lin(33, 1) = 18
Lin(34, 0) = 21: Lin(34, 1) = 19
Lin(35, 0) = 21: Lin(35, 1) = 20

Lin(36, 0) = 22: Lin(36, 1) = 23
Lin(37, 0) = 23: Lin(37, 1) = 24
Lin(38, 0) = 24: Lin(38, 1) = 25
Lin(39, 0) = 25: Lin(39, 1) = 22
Lin(40, 0) = 22: Lin(40, 1) = 24
Lin(41, 0) = 23: Lin(41, 1) = 25

Rem colors
For a = 0 To 11
Lin(a, 2) = 255
Next a
For a = 12 To 23
Lin(a, 3) = 255
Next a
For a = 24 To 35
Lin(a, 4) = 255
Next a
For a = 36 To 41
For b = 2 To 4
Lin(a, b) = 255
Next b, a

Change
Change
End Sub
