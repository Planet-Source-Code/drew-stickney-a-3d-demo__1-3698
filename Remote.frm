VERSION 5.00
Begin VB.Form Remote 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote"
   ClientHeight    =   8310
   ClientLeft      =   9450
   ClientTop       =   375
   ClientWidth     =   2550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   2550
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00808080&
      Caption         =   "E X I T"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8025
      Left            =   2040
      MaskColor       =   &H000000FF&
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.HScrollBar ScrZoom 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   9999
      TabIndex        =   18
      Top             =   7440
      Width           =   1575
   End
   Begin VB.HScrollBar ScrScale 
      Height          =   255
      LargeChange     =   20
      Left            =   240
      Max             =   5000
      Min             =   -5000
      TabIndex        =   15
      Top             =   6120
      Value           =   20
      Width           =   1575
   End
   Begin VB.HScrollBar ScrZTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   10
      Top             =   4800
      Width           =   1575
   End
   Begin VB.HScrollBar ScrYTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   9
      Top             =   4080
      Width           =   1575
   End
   Begin VB.HScrollBar ScrXTrans 
      Height          =   255
      LargeChange     =   4
      Left            =   240
      Max             =   9999
      Min             =   -9999
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.HScrollBar ScrZRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.HScrollBar ScrYRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.HScrollBar ScrXRot 
      Height          =   255
      LargeChange     =   5
      Left            =   240
      Max             =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   840
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00FF8080&
      X1              =   840
      X2              =   840
      Y1              =   7320
      Y2              =   6960
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   7320
      Y2              =   8160
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1920
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   8160
      Y2              =   6960
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FF8080&
      X1              =   840
      X2              =   1920
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label LblZm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label LblZoom 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   645
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1920
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   6840
      Y2              =   5640
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   6000
      Y2              =   6840
   End
   Begin VB.Label LblSca 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scale : 1"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   840
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   840
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF8080&
      X1              =   840
      X2              =   840
      Y1              =   5640
      Y2              =   6000
   End
   Begin VB.Label LblScale 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   615
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   1200
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      X1              =   1200
      X2              =   1200
      Y1              =   3240
      Y2              =   2880
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      X1              =   1200
      X2              =   1920
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   2880
      Y2              =   5520
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   3240
      Y2              =   5520
   End
   Begin VB.Label LblZTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label LblYTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label LblXTrans 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label LblTranslate 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Translate"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   120
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   1920
      Y1              =   2760
      Y2              =   120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      X1              =   1920
      X2              =   960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      X1              =   960
      X2              =   960
      Y1              =   120
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   2760
   End
   Begin VB.Label LblZRot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Z : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label LblYRot 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Y : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LblXRot 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "X : 0"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label LblRotate 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Remote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdExit_Click()
End
End Sub

Private Sub Form_Load()
View.Show
Remote.Show
End Sub

Private Sub ScrScale_Change()
LblSca.Caption = "Scale :" + Str$(ScrScale.Value * 0.05)
Sa = ScrScale.Value * 0.05
Change
End Sub

Private Sub ScrScale_Scroll()
LblSca.Caption = "Scale :" + Str$(ScrScale.Value * 0.05)
Sa = ScrScale.Value * 0.05
Change
End Sub

Private Sub ScrXRot_Change()
LblXRot.Caption = "X :" + Str$(ScrXRot.Value)
RotX = ScrXRot.Value
Change
End Sub

Private Sub ScrXRot_Scroll()
LblXRot.Caption = "X :" + Str$(ScrXRot.Value)
RotX = ScrXRot.Value
Change
End Sub

Private Sub ScrXTrans_Change()
LblXTrans.Caption = "X :" + Str$(ScrXTrans.Value)
TransX = ScrXTrans.Value
Change
End Sub

Private Sub ScrXTrans_Scroll()
LblXTrans.Caption = "X :" + Str$(ScrXTrans.Value)
TransX = ScrXTrans.Value
Change
End Sub

Private Sub ScrYRot_Change()
LblYRot.Caption = "Y :" + Str$(ScrYRot.Value)
RotY = ScrYRot.Value
Change
End Sub

Private Sub ScrYRot_Scroll()
LblYRot.Caption = "Y :" + Str$(ScrYRot.Value)
RotY = ScrYRot.Value
Change
End Sub

Private Sub ScrYTrans_Change()
LblYTrans.Caption = "Y :" + Str$(ScrYTrans.Value)
TransY = ScrYTrans.Value
Change
End Sub

Private Sub ScrYTrans_Scroll()
LblYTrans.Caption = "Y :" + Str$(ScrYTrans.Value)
TransY = ScrYTrans.Value
Change
End Sub

Private Sub ScrZoom_Change()
LblZm.Caption = "Zoom :" + Str$(ScrZoom.Value)
Zm = ScrZoom.Value
Change
End Sub

Private Sub ScrZoom_Scroll()
LblZm.Caption = "Zoom :" + Str$(ScrZoom.Value)
Zm = ScrZoom.Value
Change
End Sub

Private Sub ScrZRot_Change()
LblZRot.Caption = "Z :" + Str$(ScrZRot.Value)
RotZ = ScrZRot.Value
Change
End Sub

Private Sub ScrZRot_Scroll()
LblZRot.Caption = "Z :" + Str$(ScrZRot.Value)
RotZ = ScrZRot.Value
Change
End Sub

Private Sub ScrZTrans_Change()
LblZTrans.Caption = "Z :" + Str$(ScrZTrans.Value)
TransZ = ScrZTrans.Value
Change
End Sub

Private Sub ScrZTrans_Scroll()
LblZTrans.Caption = "Z :" + Str$(ScrZTrans.Value)
TransZ = ScrZTrans.Value
Change
End Sub
