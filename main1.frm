VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Enrollment System"
   ClientHeight    =   10245
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   15735
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid Grid1 
      Bindings        =   "main1.frx":0000
      Height          =   7215
      Left            =   2040
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   12726
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
      Caption         =   "Enrollment System"
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
      Height          =   450
      Left            =   8040
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   794
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblStudents"
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7180
      Left            =   0
      Picture         =   "main1.frx":0015
      ScaleHeight     =   7125
      ScaleMode       =   0  'User
      ScaleWidth      =   18.404
      TabIndex        =   10
      Top             =   2160
      Width           =   1920
      Begin VB.CommandButton cmdEditStudents 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit Students"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":E9E19F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log-In"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":E9EA69
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmdSecurity 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tools"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":E9EF5B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":E9F825
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnroll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enroll"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":E9FD2D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "main1.frx":EA006A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
         Width           =   1575
      End
   End
   Begin VB.PictureBox Header 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   2040
      Picture         =   "main1.frx":EA0934
      ScaleHeight     =   1665
      ScaleWidth      =   13545
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ENROLLMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1215
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   8895
      End
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "main1.frx":EADBDE
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Master's List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mFenroll 
         Caption         =   "E&nroll"
         Visible         =   0   'False
      End
      Begin VB.Menu mFsearch 
         Caption         =   "S&earch"
         Visible         =   0   'False
      End
      Begin VB.Menu mFsecurity 
         Caption         =   "Tools"
         Visible         =   0   'False
      End
      Begin VB.Menu mFEditStudent 
         Caption         =   "Edit Student"
         Visible         =   0   'False
      End
      Begin VB.Menu mFlogin 
         Caption         =   "Log-in"
      End
      Begin VB.Menu mFexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mHabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command22_Click()
frmLogin.Show
End Sub

Private Sub cmdEditStudents_Click()
frmEditStudents.Show
End Sub

Private Sub cmdEnroll_Click()
frmenrollment.Show

End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdHelp_Click()
frmEditStudents.Show
End Sub

Private Sub cmdLogin_Click()
Dim msg As Integer
If cmdLogin.Caption = "Log-In" Then
    frmLogin.Show
    Me.Enabled = False
Else
    msg = MsgBox("Log-Out user " & id, vbYesNo, "Message")
    If msg = vbYes Then
        cmdEnroll.Enabled = False
        
        cmdLogin.Caption = "Log-In"
        mFlogin.Caption = "Log-In"
        cmdEditStudents.Enabled = False
        cmdSearch.Enabled = False
        cmdSecurity.Enabled = False
        Grid1.Visible = False
        mFenroll.Visible = False
        mFsearch.Visible = False
        mFsecurity.Visible = False
        mFEditStudent.Visible = False

    End If
End If

End Sub

Private Sub cmdSearch_Click()
frmSearch.Show
End Sub

Private Sub cmdSecurity_Click()
frmTools.Show
End Sub



Private Sub Form_Unload(Cancel As Integer)
'end na ang program...
End
End Sub

'enroll under file menu
Private Sub mFenroll_Click()
'call cmdenroll command
cmdEnroll_Click
End Sub
'exit under file menu
Private Sub mFexit_Click()
'call cmdexit command
cmdExit_Click
End Sub
'Login under File Menu
Private Sub mFlogin_Click()
'call cmdLogin command
cmdLogin_Click
End Sub
Private Sub mfeditstudent_click()
cmdEditStudents_Click
End Sub
'Search under File Menu
Private Sub mFsearch_Click()
'call cmdsearch command
cmdSearch_Click
End Sub

Private Sub mFsecurity_Click()
cmdSecurity_Click
End Sub

Private Sub mHabout_Click()
frmAbout.Show
End Sub
