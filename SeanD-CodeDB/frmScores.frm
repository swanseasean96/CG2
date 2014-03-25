VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmScores 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   7455
   ClientTop       =   2415
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5655
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtUsers 
      DataField       =   "Username"
      DataSource      =   "adoScores"
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboUser 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Text            =   "Select User"
      Top             =   480
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc adoScores 
      Height          =   330
      Left            =   0
      Top             =   4440
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
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   1560
      Top             =   600
   End
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0000C000&
      Caption         =   "Select User"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   4815
      Left            =   0
      Top             =   0
      Width           =   5655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   5040
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Average"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 8"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   14
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 7"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   13
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 6"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 5"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 4"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 3"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 2"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Week 1"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
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
      Left            =   5400
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmScores.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
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
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
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
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboUser_Change()

Dim i As Integer    'using to manipulate index of control array

adoScores.Recordset.MoveFirst

Do Until adoScores.Recordset.Fields("Username") = cboUser.Text
    adoScores.Recordset.MoveNext
Loop

For i = 1 To 9
    If i = 9 Then
        lblScore(i).Caption = "Average " & " - " & adoScores.Recordset.Fields("Average")
    Else
        lblScore(i).Caption = "Week " & i & " - " & adoScores.Recordset.Fields("Week" & i)
    End If
Next

End Sub

Private Sub cmdPupils_Click()
frmUsers.Show
Unload Me
End Sub

Private Sub cmdSelect_Click()

Dim i As Integer    'using to manipulate index of control array

adoScores.Recordset.MoveFirst

If cboUser.Text = "Select User" Then
            MsgBox ("Please select a user")
    Else
            Do Until adoScores.Recordset.Fields("Username") = cboUser.Text
                adoScores.Recordset.MoveNext
            Loop
            
            For i = 1 To 9
                If i = 9 Then
                    lblScore(i).Caption = "Average " & " - " & adoScores.Recordset.Fields("Average")
                Else
                    lblScore(i).Caption = "Week " & i & " - " & adoScores.Recordset.Fields("Week" & i)
                End If
            Next
End If
End Sub

Private Sub cmdTests_Click()
frmTestEdit.Show
Unload Me
End Sub

Private Sub Form_Load()
Call ActiveTime

adoScores.Recordset.MoveFirst
    cboUser.AddItem (txtUsers.Text)

Do While adoScores.Recordset.EOF = False
    adoScores.Recordset.MoveNext
    cboUser.AddItem (txtUsers.Text)
    Loop
    
End Sub
Sub ActiveTime()
    lblTime.Caption = Time                  'sets the time label to the current time
    lblDate.Caption = Date                  'sets the date label to the current date
End Sub

Private Sub lblAverage_Click()

End Sub

Private Sub lblW_Click(Index As Integer)

End Sub

Private Sub lblExit_Click()
frmTeacher.Show
Unload Me
End Sub

Private Sub tmrTime_Timer()
Call ActiveTime
End Sub
