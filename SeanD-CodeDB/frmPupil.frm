VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPupil 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4545
   ClientLeft      =   6780
   ClientTop       =   4635
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdLogout 
      BackColor       =   &H0000C000&
      Caption         =   "Log Out"
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      DataSource      =   "adoTest"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cboMain 
      Height          =   405
      ItemData        =   "frmPupil.frx":0000
      Left            =   3120
      List            =   "frmPupil.frx":0002
      TabIndex        =   9
      Text            =   "Select Date:"
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H000000C0&
      Caption         =   "Ask For Help"
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdScores 
      BackColor       =   &H0000C000&
      Caption         =   "Scores"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest 
      BackColor       =   &H0000C000&
      Caption         =   "Test"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   120
      Top             =   1800
   End
   Begin MSAdodcLib.Adodc adoTest 
      Height          =   615
      Left            =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      RecordSource    =   "Year3"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoScore 
      Height          =   330
      Left            =   2040
      Top             =   720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Score"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      Height          =   4545
      Left            =   0
      Top             =   0
      Width           =   5430
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
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
      Left            =   3360
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPrevious 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Previous Score:"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblSelectDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "please select a date to continue"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1455
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
   Begin VB.Label lblYearGroup 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
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
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmPupil.frx":0004
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
      TabIndex        =   0
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
End
Attribute VB_Name = "frmPupil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub DateCheck()

If adoScore.Recordset.Fields("RecentTest") > cboMain.Text Then
    MsgBox ("please select a test which you have not yet completed")            'display message box
ElseIf adoScore.Recordset.Fields("RecentTest") = "" Then                        'clear that records field

End If

End Sub

Private Sub adoUsers_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub cmdHelp_Click()
  
frmHelp.Show                    'if student selects help then display help form and make invisible
Me.Visible = False
    
End Sub

Private Sub cmdLogout_Click()
Unload Me
Call Logout


End Sub

Private Sub cmdScores_Click()
Unload Me
frmPupilScore.Show
End Sub

Private Sub cmdTest_Click()


    If cboMain.Text = "Select Date:" Then               'if the text in the combo box is select date then make select date label visible and change the combo boxes coolour red
        lblSelectDate.Visible = True
        cboMain.BackColor = vbRed
    
    Else
        If cboMain.Text = "" Then                       'if the combo boxes text is empty iterate above ^^^
            lblSelectDate.Visible = True
            cboMain.BackColor = vbRed
        Else
            test = cboMain.Text                         'otherwise the test (global) is the text in the combo box
            frmTestMain.Show                            'display test form
            Unload Me
    End If
    End If
End Sub



Private Sub Form_Load()
Dim annum As String
Dim yeargroup As Integer
annum = "year " & group
lblUserName = user
lblYearGroup = annum

yeargroup = group
    Select Case yeargroup
    Case Is = 3
        adoTest.RecordSource = "Year3"
    Case Is = 4
        adoTest.RecordSource = "Year4"
    Case Is = 5
        adoTest.RecordSource = "Year5"
    Case Is = 6
        adoTest.RecordSource = "Year6"
    End Select
    
adoTest.Refresh

adoTest.Recordset.MoveFirst
    cboMain.AddItem (txtDate.Text)

    Do While adoTest.Recordset.EOF = False
        adoTest.Recordset.MoveNext
        cboMain.AddItem (txtDate.Text)
    Loop
    
Call ActiveTime
    
End Sub


Private Sub lblExit_Click()
Unload Me
Logout

End Sub

Sub ActiveTime()
    lblTime.Caption = Time
    lblDate.Caption = Date
End Sub





Private Sub tmrTime_Timer()
    Call ActiveTime
End Sub
