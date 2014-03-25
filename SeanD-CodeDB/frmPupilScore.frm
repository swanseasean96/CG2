VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPupilScore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   480
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H0000C000&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtDummy 
      DataField       =   "Username"
      DataSource      =   "adoPupilScore"
      Height          =   285
      Left            =   3840
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc adoPupilScore 
      Height          =   330
      Left            =   1680
      Top             =   480
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
   Begin VB.Label lblWellDone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Well Done your average is more than 75%"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image imgWellDone 
      Height          =   1740
      Left            =   480
      Picture         =   "frmPupilScore.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   2160
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
      TabIndex        =   12
      Top             =   960
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
      TabIndex        =   11
      Top             =   1320
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
      TabIndex        =   10
      Top             =   1680
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
      TabIndex        =   9
      Top             =   2040
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
      TabIndex        =   8
      Top             =   2400
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
      TabIndex        =   7
      Top             =   2760
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
      TabIndex        =   6
      Top             =   3120
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
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
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
      TabIndex        =   4
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   5040
      Y1              =   3840
      Y2              =   3840
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
      TabIndex        =   3
      Top             =   600
      Width           =   1455
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
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmPupilScore.frx":B4E4
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
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   4545
      Left            =   0
      Top             =   0
      Width           =   5430
   End
   Begin VB.Shape Shape2 
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
Attribute VB_Name = "frmPupilScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReturn_Click()
Unload Me
frmPupil.Show
End Sub

Private Sub Form_Load()
    
Dim i As Integer    'using to manipulate index of control array

    
    lblTime.Caption = Time
    lblDate.Caption = Date

adoPupilScore.Recordset.MoveFirst

Do Until adoPupilScore.Recordset.Fields("Username") = user
                adoPupilScore.Recordset.MoveNext
            Loop
            
    For i = 1 To 9
        If i = 9 Then
            lblScore(i).Caption = "Average " & " - " & adoPupilScore.Recordset.Fields("Average")
                If Val(adoPupilScore.Recordset.Fields("Average")) > 15 Then
                imgWellDone.Visible = True
                lblWellDone.Visible = True
                End If
        Else
            lblScore(i).Caption = "Week " & i & " - " & adoPupilScore.Recordset.Fields("Week" & i)
        End If
    Next
    
End Sub

Private Sub lblExit_Click()
Unload Me
Logout
End Sub

Private Sub Timer1_Timer()

    lblTime.Caption = Time
    lblDate.Caption = Date
End Sub
