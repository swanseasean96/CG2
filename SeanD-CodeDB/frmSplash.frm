VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   6500
      Left            =   3840
      Top             =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Label lblSpellBee 
      BackColor       =   &H8000000E&
      Caption         =   "Spelling Bee"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   960
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As Currency
Dim g As Currency

Sub SplashEnd()
    Unload Me                       'closes form
    frmLogin.Show                   'shows login form
End Sub

Private Sub Form_Click()
    Call SplashEnd                  'call sub routine
End Sub

Private Sub Form_Load()

Shape1.Width = 0
r = 0
g = 0
End Sub

Private Sub image1_click()
    Call SplashEnd                  'call sub routine
End Sub

Private Sub lblspellbee_click()
    Call SplashEnd                  'call sub routine
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call SplashEnd                  'call sub routine
End Sub

Private Sub Timer1_Timer()
    Call SplashEnd                  'call sub routine
End Sub

Private Sub Timer2_Timer()


Shape1.Width = Shape1.Width + 11.4999
r = r + 0.2
g = g + 0.5
Shape1.BackColor = RGB(r, g, 0)

End Sub
