VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSecurity 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   7020
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5160
      Top             =   120
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Enrolment System\dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Enrolment System\dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from login"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame frUser 
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Visible         =   0   'False
         Width           =   6615
         Begin VB.CommandButton cmdadd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add New User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1080
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Search"
            Height          =   615
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Update"
            Height          =   615
            Left            =   2400
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Delete"
            Height          =   615
            Left            =   3435
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Exit"
            Height          =   615
            Left            =   5655
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdrefresh 
            Caption         =   "Refresh"
            Height          =   615
            Left            =   4545
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Timer Timer1 
         Left            =   7200
         Top             =   0
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4048
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SECURITY SETTINGS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.Frame frAmount 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton Command9 
         Caption         =   "Exit"
         Height          =   615
         Left            =   5655
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Delete"
         Height          =   615
         Left            =   3435
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Update"
         Height          =   615
         Left            =   2340
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         Height          =   615
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Add Year Leve"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mAmount 
         Caption         =   "Amo&unt Setting"
      End
      Begin VB.Menu mSystemSec 
         Caption         =   "&System Security Settings"
      End
      Begin VB.Menu mExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click(Index As Integer)
LoginRegist.Show
frmSecurity.Adodc1.Refresh
frmSecurity.Refresh

End Sub

Private Sub Command1_Click()
frmSecurity.Hide
main.Show
End Sub

Private Sub cmdrefresh_Click()
Adodc1.Refresh
End Sub

Private Sub Command4_Click()
On Error GoTo end1
Dim msg As Integer
msg = MsgBox("Delete " & Adodc1.Recordset("userid") & " from the list?", vbYesNo, "Message")
If msg = vbYes Then
Adodc1.Recordset.Delete
End If
Exit Sub
end1:
MsgBox "No Record To Delete...", vbOKOnly, "Message"
End Sub

Private Sub Command8_Click()
Dim msg As Integer
msg = MsgBox("Delelte Row?", vbYesNo, "Message")
If msg = vbYes Then
    Adodc1.Recordset.Delete
End If
End Sub
Private Sub mAmount_Click()
frAmount.Visible = True
frUser.Visible = False
Adodc1.RecordSource = "select * from tblyearlevelpayment"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Me.Refresh
End Sub
Private Sub mSystemSec_Click()
frAmount.Visible = False
frUser.Visible = True
Adodc1.RecordSource = "select userid,usergroup from login"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
Me.Refresh
End Sub
Private Sub mExit_Click()
Unload Me
End Sub
Private Sub Command7_Click()
DataGrid1.AllowUpdate = True
End Sub
Private Sub Timer2_Timer()
On Error GoTo error_2
User = Data2.Recordset.Fields("Username")
txtSearchEngine = User
error_2:
Timer2.Enabled = False
End Sub
