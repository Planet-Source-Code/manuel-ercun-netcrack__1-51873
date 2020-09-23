VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   ControlContainer=   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   5955
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   0
      Left            =   480
      ScaleHeight     =   1215
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   840
      Width           =   50
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   1215
      Index           =   1
      Left            =   1320
      ScaleHeight     =   1215
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   840
      Width           =   50
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   50
      Index           =   2
      Left            =   840
      ScaleHeight     =   45
      ScaleWidth      =   2085
      TabIndex        =   0
      Top             =   2280
      Width           =   2085
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   3
      Left            =   5640
      Picture         =   "UserControl1.ctx":0000
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   2
      Left            =   5280
      Picture         =   "UserControl1.ctx":0A8C
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   1
      Left            =   4800
      Picture         =   "UserControl1.ctx":147F
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   0
      Left            =   2040
      Picture         =   "UserControl1.ctx":1EB3
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   3
      Left            =   5640
      Picture         =   "UserControl1.ctx":28E7
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "UserControl1.ctx":33AB
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   0
      Left            =   2040
      Picture         =   "UserControl1.ctx":3C75
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   0
      Picture         =   "UserControl1.ctx":46E3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2760
   End
   Begin VB.Image Image2 
      Height          =   255
      Index           =   1
      Left            =   4800
      Picture         =   "UserControl1.ctx":5DEE
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   270
      Index           =   2
      Left            =   5280
      Picture         =   "UserControl1.ctx":685C
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


Event Salir()
Event Minimizar()

Dim icaption As String
Dim iautosize As Boolean
Dim Ibuttons As Boolean
Public Property Get Buttons() As Boolean
Buttons = Ibuttons
End Property
Public Property Let Buttons(ByVal new_buttons As Boolean)
Ibuttons = new_buttons
If new_buttons = True Then Image3(0).Visible = True
If new_buttons = False Then Image3(0).Visible = False
PropertyChanged "Buttons"
End Property




Public Property Get Caption() As String
Caption = icaption
End Property
Public Property Let Caption(ByVal new_caption As String)
icaption = new_caption
Label1.Caption = new_caption
PropertyChanged "Caption"
End Property

Public Property Get Autosize() As Boolean
Autosize = iautosize
End Property
Public Property Let Autosize(ByVal new_autosize As Boolean)
iautosize = new_autosize
If iautosize = True Then Call UserControl_Resize
PropertyChanged "Autosize"
End Property




Private Sub image1_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
mover UserControl.Parent
End Sub



Private Sub image1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2(0).Picture <> Image2(2).Picture Then Image2(0).Picture = Image2(1).Picture
Image3(0).Picture = Image3(1).Picture
End Sub

Private Sub Image2_Click(Index As Integer)
RaiseEvent Salir
End Sub

Private Sub Image2_MouseDown(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(0).Picture = Image2(2).Picture
End Sub

Private Sub Image2_MouseMove(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(0).Picture = Image2(3).Picture
End Sub

Private Sub Image2_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(0).Picture = Image2(1).Picture
End Sub

Private Sub Image3_Click(Index As Integer)
RaiseEvent Minimizar
End Sub

Private Sub Image3_MouseDown(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(0).Picture = Image3(2).Picture
End Sub

Private Sub Image3_MouseMove(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(0).Picture = Image3(3).Picture
End Sub

Private Sub Image3_MouseUp(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(0).Picture = Image3(1).Picture
End Sub

Private Sub Picture1_MouseDown(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
mover UserControl.Parent
End Sub

Private Sub UserControl_Initialize()
Picture1(0).Move 0, Image1.Height
Picture1(1).Move (Image1.left + Image1.Width) - 50, Image1.Height
Picture1(2).Move Picture1(0).Width, (Picture1(0).left + Picture1(0).Height + Image1.Height) - 50
Label1.Move 750, 200
Image2(0).Move 4500, 200
Image3(0).Move 4400, 200
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Autosize = PropBag.ReadProperty("Autosize", False)
Caption = PropBag.ReadProperty("Caption", "Form1")
Buttons = PropBag.ReadProperty("Buttons", True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If iautosize = True Then
UserControl.Height = UserControl.Parent.Height
UserControl.Width = UserControl.Parent.Width
End If
Image1.Width = UserControl.Width
Picture1(0).Height = UserControl.ScaleHeight
Picture1(2).Width = UserControl.ScaleWidth
Picture1(2).top = UserControl.ScaleHeight - 50
Picture1(1).Height = UserControl.ScaleHeight
Picture1(1).left = UserControl.Width - 50
Image2(0).left = UserControl.Width - 500
Image3(0).left = UserControl.Width - 800
End Sub

Public Sub mover(UserControl As Form)
ReleaseCapture
SendMessage UserControl.hwnd, &HA1, 2, 0&
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Autosize", iautosize, False)
Call PropBag.WriteProperty("Caption", icaption, "Form1")
Call PropBag.WriteProperty("Buttons", Ibuttons, True)
End Sub

