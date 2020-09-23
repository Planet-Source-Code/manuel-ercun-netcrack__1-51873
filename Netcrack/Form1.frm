VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "NetCrack"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5955
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   22
      Top             =   4320
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   4560
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   600
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5775
      Begin NetCrack.UserControl2 UserControl21 
         Height          =   345
         Index           =   3
         Left            =   4440
         TabIndex        =   20
         Top             =   840
         Width           =   1095
         _extentx        =   1931
         _extenty        =   609
         caption         =   "About"
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   17
         Text            =   "3"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Index           =   3
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Index           =   2
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin NetCrack.UserControl3 UserControl31 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
      End
      Begin NetCrack.UserControl2 UserControl21 
         Height          =   345
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   840
         Width           =   1095
         _extentx        =   1931
         _extenty        =   609
         caption         =   "Options"
      End
      Begin NetCrack.UserControl2 UserControl21 
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   840
         Width           =   1095
         _extentx        =   1931
         _extenty        =   609
         caption         =   "Close"
      End
      Begin NetCrack.UserControl2 UserControl21 
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1095
         _extentx        =   1931
         _extenty        =   609
         caption         =   "Start"
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   4680
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "NetCrack"
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
         Left            =   4680
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pass:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   14
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "User:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   2920
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin NetCrack.UserControl1 UserControl11 
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7144
      Caption         =   "NetCrack"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim shel As New Class1


Private Sub Form_Load()
Me.Move (Screen.Width / 2) - Me.Width / 2, (Screen.Height / 2) - Me.Height / 2

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMinimized Then shel.ShellAdd
End Sub

Private Sub Frame1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
For i = Text1.LBound To Text1.UBound
Text1(i).BackColor = &H8000000B
Next i
Label4.Caption = ""
For i = UserControl21.LBound To UserControl21.UBound
UserControl21(i).desfocus
Next i
End Sub


Private Sub Picture1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
If button = 1 Then
shel.ShellDelete
End If
End Sub

Private Sub Text1_MouseMove(Index As Integer, button As Integer, Shift As Integer, X As Single, Y As Single)
Text1(Index).BackColor = vbWhite
Select Case Index
Case 0
Label4.Caption = "(Example-C$)"
Case 1
Label4.Caption = "(Example-\\host(IP))"
Case 2
Label4.Caption = "The user"
Case 3
Label4.Caption = "The pass"
Case 4
Label4.Caption = "The Time"
End Select
End Sub

Private Sub Text2_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = "Status NetCrack"
End Sub

Private Sub Timer1_Timer()
ser = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If Label9.Tag = "up" Then
Label9.ForeColor = vbRed
Label9.Tag = "down"
Else
Label9.ForeColor = vbBlue
Label9.Tag = "up"
End If

End Sub



Private Sub UserControl11_Minimizar()

Me.WindowState = vbMinimized
End Sub

Private Sub UserControl11_Salir()
End
End Sub


Private Sub UserControl21_click(Index As Integer)
Select Case Index
Case 0
If dic = "" Then MsgBox "Selected dictionary and user", vbCritical, "NetCrack"
If dic <> "" Then Timer1.Interval = Val(Text1(4)) * 1000: abrir dic
Case 1
Salir Form1.Text1(1) & "\" & Form1.Text1(0)
End
Case 2
Form2.Show vbModal
Case 3
Form3.Show vbModal
End Select


End Sub
