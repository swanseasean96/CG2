VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHelp 
      DataField       =   "Help"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      DataSource      =   "adoHelp"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox txtUser 
      DataField       =   "username"
      DataSource      =   "adoHelp"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   7080
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc adoHelp 
      Height          =   330
      Left            =   120
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
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
      RecordSource    =   "USERS"
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
   Begin VB.Timer tmrPlay 
      Interval        =   20
      Left            =   4920
      Top             =   5880
   End
   Begin VB.Shape Shape1 
      Height          =   6540
      Left            =   0
      Top             =   0
      Width           =   7485
   End
   Begin VB.Label lblSpellBee 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Spelling Bee"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Image imgGPS 
      Height          =   495
      Left            =   120
      Picture         =   "frmHelp.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label LblGPS 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GreenPark School"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Image Pause 
      Height          =   615
      Left            =   3480
      Picture         =   "frmHelp.frx":7E372
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   615
   End
   Begin VB.Image Stop 
      Height          =   615
      Left            =   4200
      Picture         =   "frmHelp.frx":90A57
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   615
   End
   Begin VB.Image Play 
      Height          =   615
      Left            =   2760
      Picture         =   "frmHelp.frx":A2B3F
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   615
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpHelp 
      Height          =   5640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      URL             =   "TheBeatlesHelp.mp4"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   13229
      _cy             =   9948
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Form_Load()                 'on load
tmrPlay.Enabled = True                  'enables timer
wmpHelp.Controls.Play                   'starts the video to play

adoHelp.Recordset.MoveFirst             'moves the ado to the first record

Do Until txtUser.Text = user            'start loop until the user is found (loop variable)
adoHelp.Recordset.MoveNext              'move to next record
Loop

txtHelp.Text = 1                        'change text to 1 to signify that that user has requested help
adoHelp.Recordset.Fields("Help").Value = Val(txtHelp.Text)  'sets DB value to text box value
adoHelp.Recordset.Update                'updates the DB

frmHelp.Height = 6540                   'change the height of the form to hide ado controls and text boxs

End Sub

Private Sub Pause_Click()
wmpHelp.Controls.Pause                  'pause playback
End Sub

Private Sub Play_Click()
wmpHelp.Controls.Play                   'resumes playback
End Sub

Private Sub Stop_Click()
wmpHelp.Controls.Stop                   'stops playback
Unload Me                               'closes form
frmPupil.Show                           'loads & shows pupil form
End Sub

Private Sub tmrPlay_Timer()
wmpHelp.Controls.Play                   'plays video
tmrPlay.Enabled = False                 'disable timer
End Sub

Private Sub wmpHelp_PlayStateChange(ByVal NewState As Long)

Dim answ As String
If wmpHelp.playState = wmppsStopped Then
answ = MsgBox("You Have Requested Help", vbOKOnly, "Help")

End If
End Sub
