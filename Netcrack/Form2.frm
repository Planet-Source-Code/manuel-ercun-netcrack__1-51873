VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form2"
   ScaleHeight     =   2760
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NetCrack.UserControl1 UserControl11 
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   4868
      Caption         =   "NetCrack"
      Buttons         =   0   'False
      Begin VB.Frame Frame1 
         Height          =   1995
         Left            =   120
         TabIndex        =   1
         Top             =   700
         Width           =   5120
         Begin NetCrack.UserControl2 UserControl21 
            Height          =   345
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   1440
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
            Caption         =   "Listo"
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Common"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   7
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "English"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   6
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Spanish"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   5
            Top             =   840
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Text            =   "Administrador"
            Top             =   240
            Width           =   3495
         End
         Begin NetCrack.UserControl2 UserControl21 
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   1440
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "Cancel"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dictionary:"
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
            Top             =   840
            Width           =   1290
         End
         Begin VB.Label Label1 
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
            Left            =   720
            TabIndex        =   2
            Top             =   240
            Width           =   660
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
Me.Move Form1.left + (Form1.Width / 2) - (Me.Width / 2), Form1.top + (Form1.Height / 2) - (Me.Height / 2)
dic = "spanish.txt"
End Sub

Private Sub Frame1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.BackColor = &H8000000B
For i = UserControl21.LBound To UserControl21.UBound
UserControl21(i).desfocus
Next i
End Sub





Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Text1 = "Administrador"
dic = "spanish.txt"
Case 1
Text1 = "Administrator"
dic = "england.txt"
Case 2
Text1 = "Administrador"
dic = "common.txt"

End Select
End Sub

Private Sub Text1_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.BackColor = vbWhite
End Sub

Private Sub UserControl11_Salir()
Unload Me
End Sub

Private Sub UserControl21_click(Index As Integer)
Select Case Index
Case 0
user = Text1
Unload Me
Case 1
Unload Me
End Select
End Sub
