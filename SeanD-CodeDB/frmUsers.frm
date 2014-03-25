VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUsers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Myriad Pro"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFullName 
      DataField       =   "FullName"
      DataSource      =   "adoUsers"
      Height          =   405
      Left            =   1680
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0000C000&
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add User"
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H0000C000&
      Caption         =   "Edit User"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1680
      Top             =   720
   End
   Begin VB.TextBox txtYear 
      DataField       =   "YearGroup"
      DataSource      =   "adoUsers"
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAccess 
      DataField       =   "ascceslevel"
      DataSource      =   "adoUsers"
      Height          =   405
      Left            =   1920
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "password"
      DataSource      =   "adoUsers"
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtUsername 
      DataField       =   "username"
      DataSource      =   "adoUsers"
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc adoUsers 
      Height          =   330
      Left            =   4920
      Top             =   1320
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
   Begin VB.TextBox txtUsers 
      DataField       =   "username"
      DataSource      =   "adoUsers"
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
      Left            =   6120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cboUsers 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      Height          =   5790
      Left            =   0
      Top             =   0
      Width           =   7260
   End
   Begin VB.Label lblName 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Full Name"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblAccessLevel 
      BackColor       =   &H000000FF&
      Caption         =   "You have not got permission to edit this user"
      Height          =   1095
      Left            =   4560
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblSelectUser 
      BackColor       =   &H000000FF&
      Caption         =   "Select User"
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
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
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Year Group"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblAccess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Access Level"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblUser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Username"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   6960
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmUsers.frx":0000
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
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      FillColor       =   &H0000C000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addEdit As Boolean
Dim Add As Boolean

Private Sub cmdAdd_Click()

adoUsers.Recordset.AddNew

    lblAccessLevel.Visible = False
    lblSelectUser.Visible = False
    lblUser.Visible = True
    lblPassword.Visible = True
    lblAccess.Visible = True
    lblYear.Visible = True
    txtUsername.Visible = True
    txtPassword.Visible = True
    txtAccess.Visible = True
    txtYear.Visible = True
    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdDelete.Visible = True
    
    
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    addEdit = True
    Add = True
    addEdit = True
End Sub

Private Sub cmdCancel_Click()
If addEdit = True Then
cmdAdd.Enabled = True
cmdEdit.Enabled = True
txtUsername.Visible = False
lblUser.Visible = False
txtPassword.Visible = False
lblPassword.Visible = False
txtAccess.Visible = False
lblAccess.Visible = False
txtYear.Visible = False
lblYear.Visible = False
cmdSave.Visible = False
cmdCancel.Visible = False
cmdDelete.Visible = False

Else

adoUsers.Recordset.CancelUpdate
cmdAdd.Enabled = True
cmdEdit.Enabled = True
txtUsername.Visible = False
lblUser.Visible = False
txtPassword.Visible = False
lblPassword.Visible = False
txtAccess.Visible = False
lblAccess.Visible = False
txtYear.Visible = False
lblYear.Visible = False
cmdSave.Visible = False
cmdCancel.Visible = False
cmdDelete.Visible = False
lblName.Visible = False
txtFullName.Visible = False

adoUsers.Recordset.MoveFirst

End If
End Sub

Private Sub cmdDelete_Click()
Dim answ As String
answ = MsgBox("are you sure you wish to delete this user", vbYesNo, "Delete")
If answ = "yes" Then
adoUsers.Recordset.Delete
Else
End If
End Sub

Private Sub cmdEdit_Click()
Add = False
addEdit = False
adoUsers.Recordset.MoveFirst

    Do While txtUsername.Text = cboUsers.Text = False
        adoUsers.Recordset.MoveNext
    Loop

    If cboUsers.Text = "" Then
        lblSelectUser.Visible = True
        cboUsers.BackColor = vbRed
    Else
        If txtAccess.Text = 2 Then
            lblSelectUser.Visible = False
            lblAccessLevel.Visible = True
            cboUsers.BackColor = vbWhite
        Else
            cboUsers.BackColor = vbWhite
            lblAccessLevel.Visible = False
            lblSelectUser.Visible = False
            lblUser.Visible = True
            lblPassword.Visible = True
            lblAccess.Visible = True
            lblYear.Visible = True
            txtUsername.Visible = True
            txtPassword.Visible = True
            txtAccess.Visible = True
            txtAccess.Locked = True
            txtYear.Visible = True
            txtFullName.Visible = True
            lblName.Visible = True
        End If
End If

    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdDelete.Visible = True
    addEdit = False

End Sub

Private Sub cmdSave_Click()
Dim answ As String
Dim found As Integer

found = 0

If addEdit = True Then
    Do Until adoUsers.Recordset.EOF = True Or found = 1
        If txtUsername.Text = txtUsers.Text Then
            answ = MsgBox("This username which is already in use, please try another", vbOKOnly)
            found = 1
        ElseIf adoUsers.Recordset.EOF = True Then
            adoUsers.Recordset.Update
        Else
            adoUsers.Recordset.MoveNext
        End If
    Loop
Else

    adoUsers.Recordset.Update



End If
End Sub

Private Sub Form_Load()
adoUsers.Recordset.MoveFirst
    cboUsers.AddItem (txtUsers.Text)

Do While adoUsers.Recordset.EOF = False
    adoUsers.Recordset.MoveNext
    cboUsers.AddItem (txtUsers.Text)
    Loop

lblDate.Caption = Date
lblTime.Caption = Time


End Sub

Private Sub lblExit_Click()
Unload Me
frmTeacher.Show
End Sub

Private Sub Timer1_Timer()
lblDate.Caption = Date
lblTime = Time
End Sub
