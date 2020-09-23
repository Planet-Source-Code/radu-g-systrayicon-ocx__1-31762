VERSION 5.00
Begin VB.UserControl Icon 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   480
   ToolboxBitmap   =   "Icon.ctx":0000
   Begin VB.Shape Shape1 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   23
      Shape           =   3  'Circle
      Top             =   15
      Width           =   435
   End
End
Attribute VB_Name = "Icon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************
'   Autor Radu Giurgiteanu
'   Romanian Data Soft
'   11.Feb.2002
'********************************************************************************
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
                                                    (ByVal dwMessage As Long, lpData As _
                                                                NOTIFYICONDATA) As Long

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP

Private Const WM_MOUSEMOVE = &H200       '512
Private Const WM_LBUTTONDOWN = &H201     '513
Private Const WM_LMOUSECLICK = &H202     '514
Private Const WM_LBUTTONDBLCLK = &H203   '515
Private Const WM_RBUTTONDOWN = &H204     '516
Private Const WM_RMOUSECLICK = &H205     '517

Private WithEvents X As Form
Attribute X.VB_VarHelpID = -1
Public Event MouseDown(nButton As Integer)
Public Event MouseMove(nButton As Integer)
Public Event Click()
Public Event DblClick()
'Public Event RClick()

Public Sub CreateIcon(ByVal pIcon As Picture, ByVal sToolTip As String)
    On Error GoTo Errm
    Dim Tic As NOTIFYICONDATA
    Set X = UserControl.Parent
    Tic.cbSize = Len(Tic)
    Tic.hwnd = X.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = pIcon
    Tic.szTip = sToolTip & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
    Exit Sub
Errm:
    Err.Clear
End Sub
Public Sub ChangeIcon(ByVal pIcon As Picture, ByVal sToolTip As String)
    On Error GoTo Errm
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = X.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = pIcon
    Tic.szTip = sToolTip & Chr$(0)
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
    Exit Sub
Errm:
    Err.Clear
End Sub
Public Sub DeleteIcon()
    On Error GoTo Errm
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = X.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
    Set X = Nothing
    Exit Sub
Errm:
    Err.Clear
End Sub

Private Sub UserControl_Terminate()
    DeleteIcon
End Sub

Private Sub x_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX
    Select Case X
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(Button)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(Button)
    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(Button)
    Case WM_LMOUSECLICK
        RaiseEvent Click
    'add more events here like
    'Case WM_RMOUSECLICK
    'RaiseEvent RClick
    Case Else
        'nothing
    End Select
End Sub
Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
    If frmAbout.Visible = False Then
        frmAbout.Show vbModal
    End If
End Sub
