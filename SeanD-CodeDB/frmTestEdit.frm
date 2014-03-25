VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTestEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   2790
   ClientTop       =   180
   ClientWidth     =   7935
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
   ScaleHeight     =   17837.97
   ScaleMode       =   0  'User
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H0000C000&
      Caption         =   "Select"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox cboYear 
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
      ItemData        =   "frmTestEdit.frx":0000
      Left            =   2160
      List            =   "frmTestEdit.frx":0010
      TabIndex        =   32
      Text            =   "Select Year"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description10"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   360
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   9240
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description9"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   360
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description8"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   360
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description7"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   360
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description6"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description5"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   360
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description4"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   360
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   4920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description3"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   360
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description2"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   360
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtDescription 
      DataField       =   "Description1"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox txtTest 
      DataField       =   "Date"
      DataSource      =   "adoTests"
      Height          =   405
      Left            =   6600
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word1"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   5640
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word2"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5640
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word3"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5640
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word4"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5640
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word5"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   5640
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word6"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   5640
      TabIndex        =   15
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word7"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   5640
      TabIndex        =   14
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word8"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   5640
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word9"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   5640
      TabIndex        =   12
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      DataField       =   "Word10"
      DataSource      =   "adoTests"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   5640
      TabIndex        =   11
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0000C000&
      Caption         =   "Delete"
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0000C000&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "Save"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H0000C000&
      Caption         =   "Edit Test"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000C000&
      Caption         =   "Add Test"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cboTests 
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
      Left            =   4920
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer tmr1 
      Interval        =   100
      Left            =   1560
      Top             =   600
   End
   Begin MSAdodcLib.Adodc adoTests 
      Height          =   405
      Left            =   5280
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Computing\Spelling bee\SeanD-CodeDB\Spelling Bee.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Computing\Spelling bee\SeanD-CodeDB\Spelling Bee.mdb;Persist Security Info=False"
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
   Begin VB.Shape Shape2 
      Height          =   10065
      Left            =   0
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label lblSelectYear 
      BackColor       =   &H000000FF&
      Caption         =   "Select Year"
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
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
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmTestEdit.frx":0034
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
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
      Left            =   7680
      TabIndex        =   4
      Top             =   0
      Width           =   255
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
      TabIndex        =   3
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
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblSelectTest 
      BackColor       =   &H000000FF&
      Caption         =   "Select Test"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      FillColor       =   &H0000C000&
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmTestEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Integer
Dim B As Integer
Dim C As Integer
Dim D As Integer

Sub TestDate()

adoTests.Recordset.MoveFirst                        'moves to the first record
    cboTests.AddItem (txtTest.Text)                 'adds the text in the text box as an item in the combo box

Do While adoTests.Recordset.EOF = False             'do until the record is final
    adoTests.Recordset.MoveNext                     'move to next record
    cboTests.AddItem (txtTest.Text)                 'add the text in the text box as an item in the combo box

    Loop



End Sub



Private Sub cmdAdd_Click()


adoTests.Recordset.AddNew               'add new record

    cmdSave.Visible = True              'makes command buttons visible
    cmdCancel.Visible = True
    cmdDelete.Visible = True
    
    For A = 1 To 10
    txtDescription.Item(A).Visible = True                               'make text box visible
    txtDescription.Item(A).Text = txtDescription.Item(A).DataField      'matches text to datafield
    Next A
    For B = 1 To 10
    txtWord.Item(B).Visible = True                                      'make text box visible
    txtWord.Item(B).Text = txtWord.Item(B).DataField                    'matches text to datafield
    Next B
    
cmdAdd.Enabled = False                                                  'disables buttons
cmdEdit.Enabled = False                                                 'disables buttons
End Sub

Private Sub cmdEdit_Click()

adoTests.Recordset.MoveFirst                                'move to first record

    Do While txtTest.Text = cboTests.Text = False
        adoTests.Recordset.MoveNext                         'move to next record
    Loop
        
    If cboTests.Text = "" Then
        lblSelectTest.Visible = True                        'make label visible
        cboTests.BackColor = vbRed                          'change coulour red
    Else
    For C = 1 To 10
    txtDescription.Item(C).Visible = True                   'make visivble false
    Next C
    
    For D = 1 To 10
    txtWord.Item(D).Visible = True                          'make visible true
    Next D
    
    cmdSave.Visible = True
    cmdCancel.Visible = True                                'makes visible true ^
    cmdDelete.Visible = True                                '^
End If

End Sub

Private Sub cmdSelect_Click()
Dim yeargroup As String
yeargroup = cboYear.Text

Select Case yeargroup                   'if the case is ... then adoTests.RecordSource = "Year"...
Case Is = "Year 3"
adoTests.RecordSource = "Year3"
Case Is = "Year 4"
adoTests.RecordSource = "Year4"
Case Is = "Year 5"
adoTests.RecordSource = "Year5"
Case Is = "Year 6"
adoTests.RecordSource = "Year6"
End Select

adoTests.Refresh                        'refresh the ADO

Call TestDate
End Sub

Private Sub Form_Load()

lblDate.Caption = Date                  'label = current date
lblTime.Caption = Time                  'label = current time
End Sub

Private Sub lblExit_Click()
Unload Me
frmTeacher.Show

End Sub

Private Sub tmr1_Timer()
lblDate.Caption = Date              'same as form load
lblTime.Caption = Time
End Sub
