VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmenrollment 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Enroll Student"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPay 
      Caption         =   "Payremaining Balance"
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   10920
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmenrollment.frx":0000
      Height          =   495
      Left            =   2160
      TabIndex        =   37
      Top             =   10920
      Visible         =   0   'False
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   873
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblstudents"
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
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   480
      Picture         =   "frmenrollment.frx":0015
      ScaleHeight     =   1275
      ScaleWidth      =   1275
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Enrollment process"
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   360
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
      Begin VB.CommandButton cmdOldstudent 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OLD STUDENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00FF0000&
         Picture         =   "frmenrollment.frx":214F
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000004&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmenrollment.frx":30621
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6720
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000002&
         Caption         =   "CANCEL"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frmenrollment.frx":30B13
         TabIndex        =   18
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H80000002&
         Caption         =   "REFRESH"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frmenrollment.frx":30C69
         TabIndex        =   17
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000002&
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frmenrollment.frx":30D33
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000002&
         Caption         =   "NEW STUDENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frmenrollment.frx":31096
         TabIndex        =   15
         Top             =   1560
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc payment 
      Height          =   330
      Left            =   9360
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblyearlevelpayment"
      Caption         =   "Adodc2"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Student's Information"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   7935
      Left            =   2280
      TabIndex        =   22
      Top             =   1560
      Width           =   12615
      Begin VB.ComboBox cmbSex 
         Height          =   315
         ItemData        =   "frmenrollment.frx":32DF0
         Left            =   2040
         List            =   "frmenrollment.frx":32DFA
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtstudentno 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "####-######"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtFirstName 
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtStateOrProvince 
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   5880
         Width           =   2775
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   5280
         Width           =   2775
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   315
         Left            =   8760
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtContactNo 
         Height          =   315
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   7
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox txtBalance 
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   7080
         Width           =   2755
      End
      Begin VB.ComboBox cmbYearLevel 
         Height          =   315
         ItemData        =   "frmenrollment.frx":32E0C
         Left            =   2040
         List            =   "frmenrollment.frx":32E22
         TabIndex        =   8
         Top             =   6480
         Width           =   2775
      End
      Begin VB.TextBox txtLastName 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtAddress 
         Height          =   795
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   4200
         Width           =   7815
      End
      Begin VB.TextBox txtGuardian 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   35
         Top             =   5280
         Width           =   405
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State Or Province"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   5880
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7320
         TabIndex        =   33
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4080
         TabIndex        =   32
         Top             =   1200
         Width           =   1140
      End
      Begin VB.Label Bala 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   31
         Top             =   7080
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1560
         TabIndex        =   30
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact no."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   29
         Top             =   3600
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guardian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   28
         Top             =   3000
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   27
         Top             =   4320
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   840
         TabIndex        =   26
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student no. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   25
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grade/ Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   24
         Top             =   6480
         Width           =   1290
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   23
         Top             =   2400
         Width           =   435
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENROLL A STUDENT"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   20
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmenrollment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim idno As Integer
Dim x As Integer

Private Sub cmbYearLevel_Click()

    amt = 0
    payment.RecordSource = "select * from tblPayments where yearlevel = """ & cmbYearLevel.Text & """"
    
    payment.Refresh
    With payment.Recordset
    Do Until .EOF
    
    amt = amt + .Fields("amount")
    .MoveNext
    
    Loop
    End With
  
    txtBalance.Text = amt
    
End Sub

Private Sub cmdadd_Click()

    Enble
   
    
    Adodc1.Refresh
    'condition if no record stored
    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveLast
    End If
    SetDataField
    Adodc1.Refresh
    
    If Adodc1.Recordset.RecordCount = 0 Then
        idno = 1
    Else
        Adodc1.Recordset.MoveLast
        idno = Adodc1.Recordset.RecordCount + 1
        'Adodc1.Recordset.Fields ("studentid") + 1
    End If
    Adodc1.Recordset.AddNew
    
    'increment student number for new student
    txtstudentno.Text = Format(Now(), "yyyy-") & Format(idno, "0000")
    txtLastName.SetFocus
    
End Sub

Private Sub cmdCancel_Click()
    Dim msg As Integer
    msg = MsgBox("Cancel", vbYesNo, "Message")
    If msg = vbYes Then
        KillDataField
        Adodc1.Refresh
        
        Enble2
        
        cmdRefresh_Click
        txtstudentno.Text = ""
        
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Public Sub Enble()
'para sa enable ng mga command
    Frame1.Enabled = True
    cmdadd.Enabled = False
    cmdSave.Enabled = True
    cmdExit.Enabled = False
    cmdCancel.Enabled = True
    cmdrefresh.Enabled = True
    cmdOldstudent.Enabled = False
End Sub


Private Sub cmdOldstudent_Click()
Dim srch As String
Dim lsrch, rsrch As String

    srch = InputBox("Input Student Number Or Last Name", "Search", "Student Number or Last Name")
    
    If srch = "" Then
    Exit Sub
    ElseIf Val(srch) > 0 Then
        'para maging ####-####### and ininput na number
        lsrch = Left(srch, 4)
        rsrch = Right(srch, 7)
        srch = lsrch & "-" & rsrch
        Adodc1.RecordSource = "select * from tblstudents where studentnumber = """ & srch & """"
        Adodc1.Refresh
        If (Adodc1.Recordset.RecordCount = 0) Then
            MsgBox "Record Not Found", vbOKOnly, "Message"
        Exit Sub
        Else
           
            SetDataField
            ConBalance
            cmdrefresh.Enabled = False
        End If
    Else
        Adodc1.RecordSource = "select * from tblstudents where LastName = """ & srch & """"
        Adodc1.Refresh
        
        If (Adodc1.Recordset.RecordCount = 0) Then
            MsgBox "Record Not Found", vbOKOnly, "Message"
        Exit Sub
        Else
            SetDataField
            ConBalance
            DataGrid1.Visible = True
        End If
    End If
    
    
End Sub


Private Sub cmdpay_Click()
    stdno = txtstudentno.Text
    yrlvl = cmbYearLevel.Text
    Cash.Show
End Sub

Private Sub cmdRefresh_Click()
'clear all text
    txtBalance.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAddress.Text = ""
    txtContactNo.Text = ""
    txtGuardian.Text = ""
    txtAge.Text = ""
    txtCity.Text = ""
    cmbYearLevel.Text = ""
    cmbSex.Text = ""
    txtStateOrProvince.Text = ""
    Me.Refresh

End Sub

Private Sub cmdSave_Click()
On Error GoTo errorhandle
Dim msg As Integer

msg = MsgBox("Save Record?", vbYesNo, "Message")
If msg = vbYes Then
       
        stdno = txtstudentno.Text
        yrlvl = cmbYearLevel.Text
        
        Adodc1.Recordset.Update
        KillDataField
        
        Enble2
        cmdRefresh_Click
        Adodc1.Refresh
        
        frmMain.Adodc1.Refresh
        Set frmMain.Grid1.DataSource = frmMain.Adodc1
        txtstudentno.Text = ""
        
        
    
End If
Exit Sub
errorhandle:
MsgBox "Enter a required Value", vbOKOnly, "Message"

End Sub

Private Sub Enble2()
        Frame1.Enabled = False
        cmdadd.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdExit.Enabled = True
        cmdrefresh.Enabled = False
        Frame1.Enabled = False
        cmdOldstudent.Enabled = True
        DataGrid1.Visible = False
        cmdpay.Visible = False
        
End Sub


'set datasource and datafield to textboxes and comboboxes
'para ma eliminate ang pagadd ng record kung walang record
'na nakalagay sa database
Public Sub SetDataField()

        Set txtstudentno.DataSource = Adodc1
        txtstudentno.DataField = "studentnumber"
        
        Set txtFirstName.DataSource = Adodc1
        txtFirstName.DataField = "firstname"
        
        Set txtMiddleName.DataSource = Adodc1
        txtMiddleName.DataField = "middlename"
        
        Set txtLastName.DataSource = Adodc1
        txtLastName.DataField = "lastname"
        
        Set txtGuardian.DataSource = Adodc1
        txtGuardian.DataField = "guardiansname"
        
        Set txtAddress.DataSource = Adodc1
        txtAddress.DataField = "Address"
        
        Set txtCity.DataSource = Adodc1
        txtCity.DataField = "city"
        
        Set txtStateOrProvince.DataSource = Adodc1
        txtStateOrProvince.DataField = "stateorprovince"
        
        Set txtContactNo.DataSource = Adodc1
        txtContactNo.DataField = "phonenumber"
        
        Set txtAge.DataSource = Adodc1
        txtAge.DataField = "age"
        
        Set cmbSex.DataSource = Adodc1
        cmbSex.DataField = "sex"
        
        Set cmbYearLevel.DataSource = Adodc1
        cmbYearLevel.DataField = "yearlevel"
        
        Set txtBalance.DataSource = Adodc1
        txtBalance.DataField = "balance"
        
End Sub
'set nothing to datasource and datafield
Public Sub KillDataField()
        Set txtstudentno.DataSource = Nothing
        txtstudentno.DataField = ""
        
        Set txtFirstName.DataSource = Nothing
        txtFirstName.DataField = ""
        
        Set txtMiddleName.DataSource = Nothing
        txtMiddleName.DataField = ""
        
        Set txtLastName.DataSource = Nothing
        txtLastName.DataField = ""
        
        Set txtGuardian.DataSource = Nothing
        txtGuardian.DataField = ""
        
        Set txtAddress.DataSource = Nothing
        txtAddress.DataField = ""
        
        Set txtCity.DataSource = Nothing
        txtCity.DataField = ""
        
        Set txtStateOrProvince.DataSource = Nothing
        txtStateOrProvince.DataField = ""
        
        Set txtContactNo.DataSource = Nothing
        txtContactNo.DataField = ""
        
        Set txtAge.DataSource = Nothing
        txtAge.DataField = ""
        
        Set cmbSex.DataSource = Nothing
        cmbSex.DataField = ""
        
        Set cmbYearLevel.DataSource = Nothing
        cmbYearLevel.DataField = ""
        
        Set txtBalance.DataSource = Nothing
        txtBalance.DataField = ""
        
End Sub

Private Sub DataGrid1_Click()

ConBalance
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Refresh
End Sub
Public Sub ConBalance()
    
    If Adodc1.Recordset.Fields("Balance") <> 0 Then
    
    
    cmdadd.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdrefresh.Enabled = False
    Frame1.Enabled = False
    cmdOldstudent.Enabled = False
    
    
    Else
        cmdpay.Visible = False
        Enble
        cmdrefresh.Enabled = False
    End If
End Sub

