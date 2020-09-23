VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   LinkTopic       =   "Form4"
   ScaleHeight     =   5730
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.UserControl1 UserControl11 
      Height          =   5730
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   10107
      Autosize        =   -1  'True
      Caption         =   "Edit"
      Buttons         =   0   'False
      Begin VB.Frame Frame1 
         Height          =   4935
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3855
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   3855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   240
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Move Form1.left + (Form1.Width / 2) - (Me.Width / 2), Form1.top + (Form1.Height / 2) - (Me.Height / 2)

cambio dic
End Sub

Private Sub UserControl11_Salir()
Unload Me
End Sub
