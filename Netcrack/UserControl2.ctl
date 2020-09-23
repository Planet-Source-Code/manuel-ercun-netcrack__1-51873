VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ScaleHeight     =   2775
   ScaleWidth      =   1305
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button XP"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   720
   End
   Begin VB.Image image1 
      Height          =   345
      Index           =   2
      Left            =   0
      Picture         =   "UserControl2.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image image1 
      Height          =   345
      Index           =   1
      Left            =   0
      Picture         =   "UserControl2.ctx":0CDF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1125
   End
   Begin VB.Image image1 
      Height          =   315
      Index           =   3
      Left            =   0
      Picture         =   "UserControl2.ctx":1E2F
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image image1 
      Height          =   345
      Index           =   4
      Left            =   0
      Picture         =   "UserControl2.ctx":2F4D
      Top             =   1080
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Image image1 
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "UserControl2.ctx":416C
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image image1 
      Height          =   345
      Index           =   5
      Left            =   0
      Picture         =   "UserControl2.ctx":5390
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1125
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Astyle
[Negrita] = 0
[Italic] = 1
[Both] = 2
[Normal] = 3
End Enum

Public Enum Adisabled
[Blue] = 0
[orange] = 1
End Enum

Dim ienabled As Boolean
Dim icaption As String
Dim icolor As OLE_COLOR
Dim ifocus As Adisabled
Dim istyle As Astyle
Dim ser As Boolean, sas As Boolean, sis As Boolean

Event click()


Public Property Get Focus() As Adisabled
Focus = ifocus
End Property
Public Property Let Focus(ByVal new_focus As Adisabled)
ifocus = new_focus
If ifocus = Blue Then ser = True
If ifocus = orange Then ser = False
PropertyChanged "Focus"
End Property
Public Property Get Style() As Astyle
Style = istyle
End Property
Public Property Let Style(ByVal new_sty As Astyle)
istyle = new_sty
If istyle = Negrita Then Label1.Font.Bold = True: Label1.Font.Italic = False
If istyle = Italic Then Label1.Font.Italic = True: Label1.Font.Bold = False
If istyle = Both Then Label1.Font.Bold = True: Label1.Font.Italic = True
If istyle = Normal Then Label1.Font.Bold = False: Label1.Font.Italic = False
PropertyChanged "Style"
End Property




Public Property Get Enabled() As Boolean
Enabled = ienabled
End Property
Public Property Let Enabled(ByVal new_ene As Boolean)
ienabled = new_ene
If ienabled = False Then image1(1).Picture = image1(2).Picture: sas = False
If ienabled = True Then image1(1).Picture = image1(5).Picture: sas = True
PropertyChanged "Enabled"
End Property
Public Property Get Caption() As String
Caption = icaption
End Property
Public Property Let Caption(ByVal new_cap As String)
icaption = new_cap
Label1.Caption = new_cap

Call UserControl_Resize
PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = icolor
End Property
Public Property Let ForeColor(ByVal new_fore As OLE_COLOR)
icolor = new_fore
Label1.ForeColor = new_fore
PropertyChanged "ForeColor"
End Property

Private Sub image1_Click(Index As Integer)
RaiseEvent click

End Sub

Private Sub image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
sis = True
If sas = True Then image1(1).Picture = image1(3).Picture
End Sub

Private Sub image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If sis = False Then
If sas = True Then
If ser = True Then
image1(1).Picture = image1(4).Picture
Else
image1(1).Picture = image1(0).Picture
End If
image1(1).MousePointer = 1
Label1.MousePointer = 1
End If
End If
End Sub

Private Sub image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
sis = False
If sas = True Then image1(1).Picture = image1(5).Picture
End Sub




Private Sub Label1_Click()
RaiseEvent click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
image1_MouseDown 0, 0, 0, 0, 0

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
image1_MouseMove 0, 0, 0, 0, 0
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
image1_MouseUp 0, 0, 0, 0, 0
End Sub

Private Sub UserControl_Initialize()
UserControl.Tag = "0"
image1(1).Move 0, 0
sis = False
medida
ienabled = True
icaption = "Button xp"
istyle = Normal
ifocus = orange

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Caption = PropBag.ReadProperty("Caption", "Button xp")
Enabled = PropBag.ReadProperty("Enabled", True)
ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
Style = PropBag.ReadProperty("Style", Normal)
Focus = PropBag.ReadProperty("Focus", orange)
End Sub

Private Sub UserControl_Resize()
image1(1).Width = UserControl.ScaleWidth
UserControl.Height = image1(1).Height
Label1.Left = UserControl.ScaleWidth - ((image1(1).Width / 2) + (Label1.Width / 2))
If image1(1).Width <= Label1.Width Then UserControl.Width = Label1.Width + 150
End Sub


Private Sub medida()
Dim has, hes As Long
has = image1(1).Width
hes = Label1.Width
Label1.Move ((has / 2) + 10) - (hes / 2), 60
End Sub


Public Sub desfocus()
Dim contro As Control
For Each contro In Form1.Controls
If contro.Tag <> "0" Then image1(1).Picture = image1(5).Picture
Next contro
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", icaption, "Button xp")
Call PropBag.WriteProperty("Enabled", ienabled, True)
Call PropBag.WriteProperty("ForeColor", icolor, vbBlack)
Call PropBag.WriteProperty("Style", istyle, Normal)
Call PropBag.WriteProperty("Focus", ifocus, orange)
End Sub




