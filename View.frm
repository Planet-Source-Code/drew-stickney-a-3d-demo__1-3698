VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form View 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   -14940
   ClientTop       =   240
   ClientWidth     =   12000
   DrawWidth       =   4
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin PicClip.PictureClip RotBMap 
      Left            =   0
      Top             =   1200
      _ExtentX        =   5080
      _ExtentY        =   10160
      _Version        =   327680
      Rows            =   6
      Cols            =   3
      Picture         =   "View.frx":0000
   End
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   1080
         Top             =   0
      End
      Begin VB.PictureBox RotB 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1095
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PicFrame As Integer

Private Sub Form_Load()

View.Top = 0
View.Left = 0
Remote.Show
End Sub

Private Sub Timer1_Timer()
RotB.Picture = RotBMap.GraphicCell(PicFrame)
PicFrame = PicFrame + 1
If PicFrame > 17 Then PicFrame = 0
End Sub
