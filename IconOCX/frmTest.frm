VERSION 5.00
Object = "*\A..\..\..\TEMPLO~1\t\IconOCX\IconOCX.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Form"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   555
      Index           =   2
      Left            =   2760
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   480
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Index           =   1
      Left            =   2040
      Picture         =   "frmTest.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   480
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Icon"
      Height          =   495
      Left            =   60
      TabIndex        =   7
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Top             =   1380
      Width           =   1815
   End
   Begin IconOCX.Icon Icon1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Icon"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Index           =   0
      Left            =   1320
      Picture         =   "frmTest.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   480
      Width           =   555
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Icon"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   675
      Index           =   2
      Left            =   2700
      Top             =   420
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   675
      Index           =   1
      Left            =   1980
      Top             =   420
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   675
      Index           =   0
      Left            =   1260
      Top             =   420
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Chose tooltip :"
      Height          =   195
      Left            =   1260
      TabIndex        =   5
      Top             =   1140
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Chose icon :"
      Height          =   255
      Left            =   1260
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Shp As Control
Private Sub cmdCreate_Click()
    If Shape1(0).Visible = True Then _
    Icon1.CreateIcon Picture1(0).Picture, Text1.Text
    If Shape1(1).Visible = True Then _
    Icon1.CreateIcon Picture1(1).Picture, Text1.Text
    If Shape1(2).Visible = True Then _
    Icon1.CreateIcon Picture1(2).Picture, Text1.Text
End Sub

Private Sub cmdDelete_Click()
    Icon1.DeleteIcon
End Sub

Private Sub Command1_Click()
    If Shape1(0).Visible = True Then _
    Icon1.ChangeIcon Picture1(0).Picture, Text1.Text
    If Shape1(1).Visible = True Then _
    Icon1.ChangeIcon Picture1(1).Picture, Text1.Text
    If Shape1(2).Visible = True Then _
    Icon1.ChangeIcon Picture1(2).Picture, Text1.Text
End Sub

Private Sub Picture1_Click(Index As Integer)
    Shape1(0).Visible = False
    Shape1(1).Visible = False
    Shape1(2).Visible = False
    Shape1(Index).Visible = True
End Sub
