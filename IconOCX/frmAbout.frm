VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2940
      TabIndex        =   6
      Top             =   2220
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Add, change and remove an icon in systray. Also let you specify the tooltil text. Need icon pictures to create icons."
      Height          =   855
      Left            =   960
      TabIndex        =   10
      Top             =   360
      Width           =   2475
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "RGIconOCX"
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Application :"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   75
      X2              =   3610
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   3600
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "11 . February . 2002"
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0.0"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Radu Giurgiteanu"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Author :"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'   Autor Radu Giurgiteanu
'   Romanian Data Soft
'   11.Feb.2002
'********************************************************************************
Private Sub cmdOK_Click()
    Me.Visible = False
    Unload Me
End Sub
