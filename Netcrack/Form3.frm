VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   LinkTopic       =   "Form3"
   ScaleHeight     =   1905
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NetCrack.UserControl1 UserControl11 
      Height          =   1905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   3360
      Caption         =   "NetCrack"
      Buttons         =   0   'False
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Create for ErcUn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         TabIndex        =   1
         Top             =   1200
         Width           =   2340
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form3.frx":0000
         Top             =   840
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Move Form1.left + (Form1.Width / 2) - (Me.Width / 2), Form1.top + (Form1.Height / 2) - (Me.Height / 2)

End Sub

Private Sub UserControl11_Salir()
Unload Me
End Sub
