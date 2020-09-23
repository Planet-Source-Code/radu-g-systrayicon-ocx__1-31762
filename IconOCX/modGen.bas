Attribute VB_Name = "ModGen"
'********************************************************************************
'   Autor Radu Giurgiteanu
'   Romanian Data Soft
'   11.Feb.2002
'********************************************************************************
Public Type NOTIFYICONDATA
   cbSize   As Long
   hwnd     As Long
   uID      As Long
   uFlags   As Long
   uCallbackMessage As Long
   hIcon    As Long
   szTip    As String * 64
End Type


