VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTeacher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H0000C000&
      Caption         =   "Help Selected "
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ListBox listHelp 
      Height          =   1980
      ItemData        =   "frmTeacher.frx":0000
      Left            =   2880
      List            =   "frmTeacher.frx":0002
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtPupil 
      DataField       =   "username"
      DataSource      =   "adoPupilHelp"
      Height          =   360
      Left            =   3960
      TabIndex        =   9
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtHelp 
      DataField       =   "Help"
      DataSource      =   "adoPupilHelp"
      Height          =   360
      Left            =   3120
      TabIndex        =   8
      Top             =   4320
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoPupilHelp 
      Height          =   375
      Left            =   3360
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdScores 
      BackColor       =   &H0000C000&
      Caption         =   "Scores"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdTests 
      BackColor       =   &H0000C000&
      Caption         =   "Tests"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdPupils 
      BackColor       =   &H0000C000&
      Caption         =   "Pupils"
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   120
      Top             =   1800
   End
   Begin VB.Shape Shape2 
      Height          =   4845
      Left            =   0
      Top             =   0
      Width           =   5400
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblSpellBee 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Spelling Bee"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmTeacher.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
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
      Left            =   -240
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "frmTeacher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub RetrieveInfo()
Dim Username As String                      'sets "username" as a variable
Dim password As String                      'sets "password" as a variable
Dim yeargroup As String                     'sets "yeargroup" as a variable
Dim helpNeeded As String
Dim helped As String


Username = frmLogin.txtUserCheck            'sets "username" to = the username on the login form
password = frmLogin.txtPassCheck            'sets "password" to = the username on the login form
yeargroup = frmLogin.txtYearGroup           'sets "yeargroup" to = the username on the login form

lblUserName.Caption = Username              'the username label caption = the username variable

End Sub


Sub ActiveTime()
    lblTime.Caption = Time                  'sets the time label to the current time
    lblDate.Caption = Date                  'sets the date label to the current date
End Sub

Private Sub cmdHelp_Click()


adoPupilHelp.Recordset.MoveFirst

helped = listHelp.List(listHelp.ListIndex)

If listHelp.List(listHelp.ListIndex) = "" Then
answ = MsgBox("There are no pupils requiring help at the moment", vbOKOnly, Error)
Else
Do Until txtPupil.Text = helped
adoPupilHelp.Recordset.MoveNext
Loop
txtHelp.Text = 0
adoPupilHelp.Recordset.Update
listHelp.RemoveItem (listHelp.ListIndex)
adoPupilHelp.Recordset.MoveFirst
End If

End Sub

Private Sub cmdPupils_Click()
frmUsers.Show
Unload Me
End Sub

Private Sub cmdScores_Click()
frmScores.Show
Unload Me
End Sub

Private Sub cmdTests_Click()
frmTestEdit.Show
Unload Me
End Sub

Private Sub Form_Load()
Call ActiveTime
Call RetrieveInfo                           'calls the RetrieveInfo sub routine

adoPupilHelp.Refresh

adoPupilHelp.Recordset.MoveFirst
Do Until adoPupilHelp.Recordset.EOF = True
If txtHelp.Text = 0 Then

adoPupilHelp.Recordset.MoveNext
Else
helpNeeded = txtPupil.Text
listHelp.AddItem (helpNeeded)
adoPupilHelp.Recordset.MoveNext
End If
Loop
End Sub

Private Sub lblExit_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub tmrTime_Timer()
    Call ActiveTime                         'call ActiveTime sub routine
End Sub

