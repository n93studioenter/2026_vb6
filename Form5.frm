VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   LinkTopic       =   "Form5"
   ScaleHeight     =   825
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public msg As String
Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height - 400
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
