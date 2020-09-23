VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200


Dim sn As NOTIFYICONDATA

Private Sub Class_Initialize()
sn.cbSize = Len(sn)
sn.hIcon = Form1.Picture1.Picture
sn.hwnd = Form1.Picture1.hwnd
sn.uID = 1
sn.szTip = "NetCrack for ErcUn"
sn.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
sn.uCallbackMessage = WM_MOUSEMOVE
End Sub


Public Function ShellAdd()
Shell_NotifyIcon NIM_ADD, sn
Form1.Hide
End Function

Public Function ShellDelete()
Shell_NotifyIcon NIM_DELETE, sn
Form1.WindowState = vbNormal
Form1.Show
End Function




