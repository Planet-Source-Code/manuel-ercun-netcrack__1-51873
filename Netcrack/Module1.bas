Attribute VB_Name = "Module1"
Option Explicit

Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  lpLocalName As String
  lpRemoteName As String
  lpComment As String
  lpProvider As String
End Type



Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Const RESOURCETYPE_ANY = &H0&


Public i As Long
Dim j As Long
Public user As String, dic As String
Public ser As Boolean

Public Sub abrir(paht As String)

Dim a As NETRESOURCE
Dim res As Long
Dim g As String

a.lpLocalName = vbNullString
a.lpProvider = vbNullString
a.dwType = RESOURCETYPE_ANY
a.lpRemoteName = Form1.Text1(1) & "\" & Form1.Text1(0)
Open paht For Input As #1
Form1.UserControl31.Min = 0
Form1.UserControl31.Max = LOF(1)
Do
j = j + 1
Line Input #1, g
res = WNetAddConnection2(a, g, user, 0)
Form1.Text1(2) = user
Form1.Text1(3) = g
If res <> 0 Then
Form1.Text2 = LastErrorApi(res)
End If
If res = 0 Then
Form1.Text1(2) = user
Form1.Text1(3) = g
Form1.Text2 = "YEAHHHH!!!!" & LastErrorApi(res)
Exit Do
End If
Form1.UserControl31.Value = j
ser = False
Form1.Timer1.Enabled = True
Do
DoEvents
Loop Until ser = True

Loop Until EOF(1)
Close #1
Salir a.lpRemoteName

End Sub

Public Sub Salir(cus As String)
WNetCancelConnection2 cus, 0&, 1
End Sub
Public Function LastErrorApi(Errordll As Long) As String
Dim res&
Dim s$
s = String(255, vbNullChar)
res = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, Errordll, 0, s, Len(s), 255)
If res <> 0 Then
LastErrorApi = left(s, res)
End If
End Function



