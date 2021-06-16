VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Search"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Report 
      BackColor       =   &H00C0C000&
      Caption         =   "Report"
      Height          =   2295
      Left            =   360
      TabIndex        =   15
      Top             =   3480
      Width           =   3255
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
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
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   2160
         Picture         =   "frmSearch.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSearch.frx":04F2
      Height          =   5775
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12632319
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
            LCID            =   16393
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
            LCID            =   16393
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11400
      Top             =   240
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
      RecordSource    =   "select * from tblStudents"
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
   Begin VB.CommandButton cmdpay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pay Remaining Balance"
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
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Height          =   2295
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   3255
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmSearch.frx":0507
         Left            =   240
         List            =   "frmSearch.frx":051D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search For Last Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1905
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search For Year Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1920
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE STUDENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblBalance 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   3540
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAmountPaid 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   3540
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remaining Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line Line9 
      X1              =   240
      X2              =   12960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   13080
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      X1              =   12960
      X2              =   12960
      Y1              =   7080
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   13080
      X2              =   13080
      Y1              =   6960
      Y2              =   240
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   240
      Y1              =   5520
      Y2              =   5640
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   12960
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   13080
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   6960
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   7080
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Search Student"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3923
      TabIndex        =   7
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEdit_Click()
With frmEditStudents
.cmdEdit.Caption = "Update"
.Frame1.Enabled = True
.Adodc1.RecordSource = "select * from tblstudents where studentnumber = """ & Adodc1.Recordset.Fields("studentnumber") & """"
.Adodc1.Refresh
Set .DataGrid1.DataSource = .Adodc1
.Show

End With
cmdRefresh_Click
End Sub

Private Sub cmdpay_Click()
stdno = Adodc1.Recordset.Fields("studentnumber")
yrlvl = Adodc1.Recordset.Fields("yearlevel")

Cash.Show
End Sub

Private Sub cmdRefresh_Click()
lblBalance.Caption = ""
lblAmountPaid.Caption = ""
cmdEdit.Enabled = False
End Sub

Private Sub cmdSearch_Click()
Adodc1.RecordSource = "select * from tblStudents where lastname = """ & txtSearch.Text & """"
Adodc1.Refresh
cmdRefresh_Click

End Sub

Private Sub Combo1_Click()
cmdRefresh_Click
Adodc1.RecordSource = "select * from tblStudents where YearLevel = """ & Combo1.Text & """"
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
Set DataStudentList.DataSource = Adodc1
DataStudentList.Show
cmdRefresh_Click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub DataGrid1_Click()
On Error GoTo end1
lblBalance.Caption = Adodc1.Recordset.Fields("balance")
lblAmountPaid.Caption = Adodc1.Recordset.Fields("payments")
stdno = Adodc1.Recordset.Fields("studentnumber")
cmdEdit.Enabled = True
end1:
End Sub

Private Sub lblAmountPaid_Change()
If lblBalance.Caption = "" And lblAmountPaid.Caption = "" Then
    cmdpay.Enabled = False
Else
    cmdpay.Enabled = True
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
End If
End Sub
