VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditStudents 
   Caption         =   "Edit Student"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   15090
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   285
      Left            =   3360
      TabIndex        =   33
      Top             =   10200
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   360
      TabIndex        =   32
      Top             =   10200
      Width           =   2775
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   12600
      TabIndex        =   31
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
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
      Left            =   13920
      TabIndex        =   30
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   11280
      TabIndex        =   29
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
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
      Left            =   9960
      TabIndex        =   28
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Student's Information"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   15015
      Begin VB.TextBox txtAge 
         DataField       =   "Age"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtGuardian 
         DataField       =   "GuardiansName"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   2640
         Width           =   2775
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   795
         Left            =   2160
         TabIndex        =   12
         Top             =   3600
         Width           =   8895
      End
      Begin VB.TextBox txtLastName 
         DataField       =   "LastName"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbYearLevel 
         DataField       =   "YearLevel"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmEditStudents.frx":0000
         Left            =   2160
         List            =   "frmEditStudents.frx":0016
         TabIndex        =   10
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox txtBalance 
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   6000
         Width           =   2775
      End
      Begin VB.TextBox txtContactNo 
         DataField       =   "PhoneNumber"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   3120
         Width           =   2775
      End
      Begin VB.TextBox txtMiddleName 
         DataField       =   "MiddleName"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   9120
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   4560
         Width           =   2775
      End
      Begin VB.TextBox txtStateOrProvince 
         DataField       =   "StateOrProvince"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   5040
         Width           =   2775
      End
      Begin VB.TextBox txtFirstName 
         DataField       =   "FirstName"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   5520
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtstudentno 
         DataField       =   "StudentNumber"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "####-######"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbSex 
         DataField       =   "sex"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "frmEditStudents.frx":0050
         Left            =   2160
         List            =   "frmEditStudents.frx":005A
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
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
         TabIndex        =   27
         Top             =   2160
         Width           =   435
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
         Left            =   720
         TabIndex        =   26
         Top             =   5520
         Width           =   1290
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
         Left            =   840
         TabIndex        =   25
         Top             =   720
         Width           =   1245
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
         Height          =   240
         Left            =   960
         TabIndex        =   24
         Top             =   1200
         Width           =   1125
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
         TabIndex        =   23
         Top             =   3840
         Width           =   885
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
         Left            =   1080
         TabIndex        =   22
         Top             =   2640
         Width           =   960
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
         TabIndex        =   21
         Top             =   3120
         Width           =   1170
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
         TabIndex        =   20
         Top             =   1680
         Width           =   405
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
         TabIndex        =   19
         Top             =   6000
         Width           =   870
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
         Height          =   240
         Left            =   4320
         TabIndex        =   18
         Top             =   1200
         Width           =   1140
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
         Height          =   240
         Left            =   7680
         TabIndex        =   17
         Top             =   1200
         Width           =   1395
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
         Left            =   240
         TabIndex        =   16
         Top             =   5040
         Width           =   1830
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
         TabIndex        =   15
         Top             =   4560
         Width           =   405
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEditStudents.frx":006C
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   4683
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
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   9960
      Width           =   2775
   End
End
Attribute VB_Name = "frmEditStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Adodc1.Refresh
End Sub

Private Sub cmdDelete_Click()
Dim msg As String
msg = MsgBox("Delete Record?", vbYesNo, "Message")
If msg = vbYes Then
Adodc1.Recordset.Delete
End If
End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "Update" Then
    Adodc1.Recordset.Update
    Frame1.Enabled = False
    frmMain.Adodc1.Refresh
    Set frmMain.Grid1.DataSource = frmMain.Adodc1
    cmdEdit.Caption = "Edit"
    Else
    Frame1.Enabled = True
    cmdEdit.Caption = "Update"
    End If
End Sub

Private Sub cmdSearch_Click()
    Adodc1.RecordSource = "select * from tblStudents where lastname = """ & txtSearch.Text & """"
    Adodc1.Refresh
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdSearch_Click
End If
End Sub


