VERSION 5.00
Begin VB.UserControl UserControl3 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   ScaleHeight     =   645
   ScaleWidth      =   4200
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "UserControl3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type Rect
     left As Long
     top As Long
     right As Long
     button As Long
End Type

Public Enum iborde
[Flat] = 0
[Fixed] = 1
End Enum

Dim cor As Rect

Dim imin, imax, ivalue As Long
Dim iBackColor As OLE_COLOR
Dim iFillColor As OLE_COLOR
Dim iForeColor As OLE_COLOR
Dim iAppearance As iborde
Dim esc As String

Public Property Get Appearance() As iborde
Appearance = iAppearance
End Property
Public Property Let Appearance(ByVal new_Appearance As iborde)
iAppearance = new_Appearance
Picture1.Appearance = new_Appearance
PropertyChanged "Appearance"
End Property



Public Property Get Min() As Long
Min = imin
End Property
Public Property Let Min(ByVal new_min As Long)
imin = new_min
If imin > imax Then imin = imax
If imin > ivalue Then imin = ivalue
PropertyChanged "Min"
End Property

Public Property Get Max() As Long
Max = imax
End Property
Public Property Let Max(ByVal new_max As Long)
imax = new_max
If imax < imin Then imax = imin
If imax < ivalue Then imax = ivalue
PropertyChanged "Max"
End Property

Public Property Get Value() As Long
Value = ivalue
End Property
Public Property Let Value(ByVal new_value As Long)
ivalue = new_value
If ivalue < imin Then ivalue = imin
If ivalue > imax Then ivalue = imax

Call Progress(Picture1, ivalue)

PropertyChanged "Value"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = iBackColor
End Property
Public Property Let BackColor(ByVal new_color As OLE_COLOR)
iBackColor = new_color
Picture1.BackColor = new_color
PropertyChanged "BackColor"
End Property


Public Property Get FillColor() As OLE_COLOR
FillColor = iFillColor
End Property
Public Property Let FillColor(ByVal new_FillColor As OLE_COLOR)
iFillColor = new_FillColor
Picture1.FillColor = new_FillColor
PropertyChanged "FillColor"
End Property


Public Property Get ForeColor() As OLE_COLOR
ForeColor = iForeColor
End Property
Public Property Let ForeColor(ByVal new_ForeColor As OLE_COLOR)
iForeColor = new_ForeColor
Picture1.ForeColor = new_ForeColor
PropertyChanged "ForeColor"
End Property



Public Sub Progress(pic As PictureBox, ByRef por As Long)
On Error Resume Next
Dim pors As Long
pors = Screen.TwipsPerPixelX

cor.left = pors
cor.top = pors
cor.right = pic.ScaleWidth - pors
cor.button = pic.ScaleHeight - pors


pic.DrawMode = 13

esc = CStr(Format(por / imax, "0.0%"))

pic.Line (cor.left, cor.top)-(cor.right, cor.button), pic.BackColor, BF
pic.CurrentX = (pic.ScaleWidth - pic.TextWidth(esc)) / 2
pic.CurrentY = (pic.ScaleHeight - pic.TextHeight(esc)) / 2
pic.Print esc

If por > 0 Then
pic.DrawMode = 7
pic.Line (cor.left, cor.top)-((cor.right / imax) * por, cor.button), pic.FillColor, BF
pic.Line (cor.left, cor.top)-((cor.right / imax) * por, cor.button), pic.BackColor, BF
End If


End Sub

Private Sub UserControl_Initialize()
imin = 0
imax = 100

ivalue = 0

iForeColor = vbBlue
iFillColor = vbBlue
iBackColor = vbWhite
Picture1.Move 0, 0
iAppearance = 0
Progress Picture1, 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
BackColor = .ReadProperty("BackColor", vbWhite)
FillColor = .ReadProperty("FillColor", vbBlue)
ForeColor = .ReadProperty("ForeColor", vbBlue)
Min = .ReadProperty("Min", 0)
Max = .ReadProperty("Max", 100)
Value = .ReadProperty("Value", 0)
BorderStyle = .ReadProperty("Appearance", 0)
End With
End Sub


Private Sub UserControl_Resize()
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("BackColor", iBackColor, vbWhite)
   Call .WriteProperty("ForeColor", iForeColor, vbBlue)
   Call .WriteProperty("FillColor", iFillColor, vbBlue)
   Call .WriteProperty("Min", imin, 0)
   Call .WriteProperty("Max", imax, 100)
   Call .WriteProperty("Value", ivalue, 0)
  Call .WriteProperty("Appearance", iAppearance, 0)
End With
End Sub



