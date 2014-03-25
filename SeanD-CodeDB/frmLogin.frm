VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   5235
   ClientLeft      =   5145
   ClientTop       =   5040
   ClientWidth     =   5805
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtYearGroup 
      DataField       =   "YearGroup"
      DataSource      =   "adoLogin"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtAccess 
      DataField       =   "ascceslevel"
      DataSource      =   "adoLogin"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtPassCheck 
      DataField       =   "password"
      DataSource      =   "adoLogin"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtUserCheck 
      DataField       =   "username"
      DataSource      =   "adoLogin"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   270
   End
   Begin MSAdodcLib.Adodc adoLogin 
      Height          =   495
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Spelling Bee.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Spelling Bee.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "USERS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H0000C000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3600
      Width           =   2895
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   5235
      Left            =   0
      Top             =   0
      Width           =   5805
   End
   Begin VB.Label lblSpellBee 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Spelling Bee"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      FillColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblPass 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      X1              =   120
      X2              =   5640
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label LblGPS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GreenPark School"
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
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      Picture         =   "frmLogin.frx":22DB7
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
    txtUser = ""                    'clears text box
    txtPass = ""                    'clears text box
End Sub

Sub Access()                            'sub routine

     If txtAccess.Text = "1" Then       'if the acces level is one then
     frmPupil.Show                      'and load pupil welcome form
     Unload Me                          'close
     Else
     frmTeacher.Show                    'and load teacher welcome form
     Unload Me                          'close
     End If
     
        

End Sub
    
    Sub saving()

    End Sub
    
Sub Login()                                               'sub routine

    Dim answ As String
    
        If txtUserCheck.Text = txtUser.Text Then          'if usernames match
            If txtPassCheck.Text = txtPass.Text Then      'check passwords
                user = txtUser.Text
                password = txtPass.Text
                group = txtYearGroup.Text
                Call Access                               'call access sub routine
                Else
                answ = MsgBox("Password or Username is incorrect please try again", vbOKOnly, "Error")
            End If
        Else
           Call NextRecord                                'otherwise call nextrecord sub routine
        End If
    
End Sub
    
Sub NextRecord()
    
    Dim answ As String
    
    If adoLogin.Recordset.EOF = True Then                                                           'if the ADO is at end of file then
        
        
        answ = MsgBox("Password or Username is incorrect please try again", vbOKOnly, "Error")      'display message box
        txtUser.Text = ""                                                                           'clear username text box
        txtPass.Text = ""                                                                           'clear password text box
        adoLogin.Recordset.MoveFirst                                                                'move first record
        
    Else
    
        adoLogin.Recordset.MoveNext                                                                     'move ADO to next record
        Call Login                                                                                      'call login sub routine
    
    End If
End Sub
    
    
    
Private Sub cmdLogin_Click()
   
    adoLogin.Recordset.MoveFirst                'move ADO to first record
        Call Login                              'call login sub routine
    
End Sub



Private Sub lblExit_Click()
    Unload Me                                   'close form
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then                           'if key pressed is 13 on ascii code then
       adoLogin.Recordset.MoveFirst                 'move ADO to first record
       Call Login                                   'call login sub routine
    End If
End Sub
