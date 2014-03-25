VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTestMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   930
   ClientTop       =   555
   ClientWidth     =   9165
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
   ScaleHeight     =   12474.94
   ScaleMode       =   0  'User
   ScaleWidth      =   9216.487
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtscore 
      DataField       =   "Week1"
      DataSource      =   "adoScores"
      Height          =   405
      Left            =   5640
      TabIndex        =   39
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtusername 
      DataField       =   "Username"
      DataSource      =   "adoScores"
      Height          =   405
      Left            =   6480
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoScores 
      Height          =   330
      Left            =   7800
      Top             =   9000
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
      Connect         =   $"frmTesting.frx":0000
      OLEDBString     =   $"frmTesting.frx":00AC
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Score"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Pro"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word10"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   10
      Left            =   7920
      TabIndex        =   37
      Top             =   7560
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.CommandButton cmdComplete 
      BackColor       =   &H0000C000&
      Caption         =   "c o m p l e t e"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   10
      Left            =   5280
      TabIndex        =   32
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   9
      Left            =   5280
      TabIndex        =   31
      Top             =   6840
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   8
      Left            =   5280
      TabIndex        =   30
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   7
      Left            =   5280
      TabIndex        =   29
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   6
      Left            =   5280
      TabIndex        =   28
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   5
      Left            =   5280
      TabIndex        =   27
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   4
      Left            =   5280
      TabIndex        =   26
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   3
      Left            =   5280
      TabIndex        =   25
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      Height          =   615
      Index           =   2
      Left            =   5280
      TabIndex        =   24
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtWord 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      Left            =   5280
      TabIndex        =   23
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txtTestDate 
      DataField       =   "Date"
      DataSource      =   "adoMainTest"
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word9"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   9
      Left            =   7920
      TabIndex        =   20
      Top             =   6840
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word8"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   8
      Left            =   7920
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word7"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   7
      Left            =   7920
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word6"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   6
      Left            =   7920
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word5"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   5
      Left            =   7920
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word4"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   4
      Left            =   7920
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word3"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   3
      Left            =   7920
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word2"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   2
      Left            =   7920
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   994
   End
   Begin VB.TextBox txtWordCheck 
      DataField       =   "Word1"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   1
      Left            =   7920
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   994
   End
   Begin MSAdodcLib.Adodc adoMainTest 
      Height          =   855
      Left            =   -120
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1508
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Spelling Bee.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Spelling Bee.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Year3"
      Caption         =   "adoMainTest"
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
      Height          =   9375
      Left            =   0
      Top             =   0
      Width           =   7455
   End
   Begin VB.Label lblTestDate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Date Of Test"
      Height          =   495
      Left            =   720
      TabIndex        =   36
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label cmdSubmit 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Submit"
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
      Left            =   5760
      TabIndex        =   35
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Score "
      Height          =   495
      Left            =   2400
      TabIndex        =   33
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "extend to right to see hidden text boxes with answers"
      Height          =   615
      Left            =   3600
      TabIndex        =   22
      Top             =   8760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 10"
      DataField       =   "Description10"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   7560
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 9"
      DataField       =   "Description9"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   6840
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 8"
      DataField       =   "Description8"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 7"
      DataField       =   "Description7"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 6"
      DataField       =   "Description6"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 5"
      DataField       =   "Description5"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 4"
      DataField       =   "Description4"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 3"
      DataField       =   "Description3"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 2"
      DataField       =   "Description2"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description 1"
      DataField       =   "Description1"
      DataSource      =   "adoMainTest"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
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
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmTesting.frx":0158
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblExit 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Enabled         =   0   'False
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
      Left            =   7200
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
      Left            =   -240
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmTestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ORANGE = &H80FF&
Dim n As Integer
Dim total As Integer
Dim found As Boolean


Sub Load()
Dim year As Integer
year = group

     Select Case year
        Case Is = 3
            adoMainTest.RecordSource = "Year3"
        Case Is = 4
            adoMainTest.RecordSource = "Year4"
        Case Is = 5
            adoMainTest.RecordSource = "Year5"
        Case Is = 6
            adoMainTest.RecordSource = "Year6"
    End Select


adoMainTest.Refresh
adoMainTest.Recordset.MoveFirst
    
    If txtTestDate.Text = test Then
    
    Else
        Do
            If adoMainTest.Recordset.EOF = True Then
            adoMainTest.Recordset.MoveFirst
            Else
            adoMainTest.Recordset.MoveNext
            End If
        Loop Until txtTestDate.Text = test
        
    End If
    

End Sub

Sub Characters()                        'minor error :)
Dim C As Integer
Dim character As Integer
Dim length As Integer
Dim percentage As Integer

    For C = 1 To Len(txtWordCheck(n).Text)
        If Mid$(txtWord(n).Text, C, 1) = Mid$(txtWordCheck(n).Text, C, 1) Then
            character = character + 1
        Else
            character = character
        End If
        
    Next C
    
    If character = 0 Then
        found = False
    Else
        length = Len(txtWordCheck(n).Text)
        percentage = (length / character) * 100
        
        If percentage > 74 Then
            total = total + 1
            txtWord(n).BackColor = ORANGE
            found = True
        End If
                            '''''''''LINE 69 YEEHAA
    End If

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function Mark() As Integer
Dim U As Integer
Dim answ As String
Dim C As Integer
Dim character As Integer
Dim length As Integer
Dim percentage As Integer

   
    total = 0
   
    For n = 1 To 10
        If txtWord(n).Text = txtWordCheck(n).Text Then
             txtWord(n).BackColor = vbGreen
             total = total + 2
             found = True
             
        Else
          
          
            Call Characters
                 
             
        End If
          
                    
          
             If found = False Then
             txtWord(n).BackColor = vbRed
             total = total
             End If
    Next n
    
    Mark = total
    
    cmdSubmit.Caption = "Return"
    lblScore.Caption = "score " & total

    score = total

    lblScore.Visible = True
     

End Function







Private Sub cmdComplete_Click()

    txtWord(1).Text = txtWordCheck(1).Text
    txtWord(2).Text = txtWordCheck(2).Text
    txtWord(3).Text = txtWordCheck(3).Text
    txtWord(4).Text = txtWordCheck(4).Text
    txtWord(5).Text = txtWordCheck(5).Text
    txtWord(6).Text = txtWordCheck(6).Text
    txtWord(7).Text = txtWordCheck(7).Text
    txtWord(8).Text = txtWordCheck(8).Text
    txtWord(9).Text = txtWordCheck(9).Text
    txtWord(10).Text = txtWordCheck(10).Text


End Sub


Private Sub cmdsubmit_click()

If cmdSubmit.Caption = "Return" Then
        Unload Me
        frmPupil.Show
        frmPupil.lblPrevious.Visible = True
        frmPupil.lblScore.Visible = True
        frmPupil.lblScore.Caption = total
        Unload Me
Else
    score = Mark
    
    Call Datasave
    
End If
                    
                    adoScores.Recordset.Fields("RecentTest") = test
                    adoScores.Recordset.Update


End Sub



    

        
        

Sub Datasave()

Dim q As Integer
Dim averageloop As Integer
Dim averagecount As Integer
Dim average As Integer
Dim totalaverage As Integer
Dim i As Integer
Dim w As Integer
    averagecount = 0

        q = 1
        i = 1
        Do Until txtUsername.Text = user
            adoScores.Recordset.MoveNext
        Loop
        
        Do Until txtscore.Text = "" Or txtscore.DataField = "Week8"
              
                    txtscore.DataField = "Week" & q
                    q = q + 1
            
        Loop
        
        If txtscore.DataField = "Week8" Then
                adoScores.Recordset.Fields("Week" & 8) = total
                
                For averageloop = 1 To 8
                    averagecount = averagecount + Val(txtscore.Text)
                    txtscore.DataField = "Week" & i
                     i = i + 1
                Next
                    average = averagecount / 8
                adoScores.Recordset.Fields("Average") = average
                adoScores.Recordset.Update

'                For w = 1 To 8
'                    adoScores.Recordset.Fields("Week" & w) = ""
'                    adoScores.Recordset.Update
'                Next
        Else
                
                adoScores.Refresh
                
                Do Until txtUsername.Text = user
                    adoScores.Recordset.MoveNext
                Loop
                
                txtscore.Text = total
                adoScores.Recordset.Fields("Week" & q) = total
                adoScores.Recordset.Update
            End If
                
            

            


End Sub

Sub UnloadForm()

Unload Me

End Sub

Private Sub Form_Load()
lblTestDate.Caption = test
End Sub

Private Sub lblExit_Click()
    Unload Me
    Unload frmPupil
    Logout

End Sub

Private Sub Text1_Change()

End Sub

