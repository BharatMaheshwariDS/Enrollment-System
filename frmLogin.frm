VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   7695
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4546.461
   ScaleMode       =   0  'User
   ScaleWidth      =   13534.91
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   6000
      Picture         =   "frmLogin.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   2955
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc adologin 
      Height          =   330
      Left            =   1080
      Top             =   480
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
      RecordSource    =   "login"
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
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   6720
      TabIndex        =   1
      Top             =   3600
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6720
      TabIndex        =   4
      Top             =   4800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   390
      Left            =   8040
      TabIndex        =   5
      Top             =   4800
      Width           =   1020
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4200
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C000&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5625
      TabIndex        =   0
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C0C000&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   5625
      TabIndex        =   2
      Top             =   4260
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'enable main form
    frmMain.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()

    'backup para kung walang nakalagay na record sa database
    If txtusername.Text = "Admin" And txtpassword.Text = "Admin" Then
        id = txtusername.Text
        'enable nya yung nasa main form kung accepted na ang userlogin
        With frmMain
                .Enabled = True
                .cmdEnroll.Enabled = True
                .cmdEditStudents.Enabled = True
                .cmdLogin.Caption = "Log-Out"
                .mFlogin.Caption = "Log-Out"
                .cmdSearch.Enabled = True
                .cmdSecurity.Enabled = True
                .mFenroll.Visible = True
                .mFsearch.Visible = True
                .mFsecurity.Visible = True
                .cmdSecurity.Enabled = True
                .Grid1.Visible = True
        End With
        Unload Me
        Exit Sub
    Else
        'nagrefresh para bumalik sa bigining file ng database
        adologin.Refresh
        'check for stored username and password
        Do Until adologin.Recordset.EOF 'hangat nasa End Of FIle(EOF) maglulup sya
        
            With adologin
            'condition kung equal nga sa user at password
            If .Recordset("userId") = txtusername.Text And _
                .Recordset("password") = txtpassword.Text Then
                frmMain.Enabled = True
                id = txtusername.Text
                
                'enable nya yung nasa main form kung accepted na ang userlogin
                With frmMain
                    
                    .cmdEnroll.Enabled = True
                    .cmdEditStudents.Enabled = True
                    .cmdLogin.Caption = "Log-Out"
                    .mFlogin.Caption = "Log-Out"
                    .cmdSearch.Enabled = True
                    .cmdSecurity.Enabled = True
                    .mFenroll.Visible = True
                    .mFsearch.Visible = True
                    .mFsecurity.Visible = True
                    .mFEditStudent.Visible = True
                    .Grid1.Visible = True
                    'condition kung admin or guest ang user
                    If adologin.Recordset.Fields("usergroup") = "GUEST" Then
                        .cmdSecurity.Enabled = False
                    Else
                    End If
                End With
                Unload Me
                Exit Sub ' exit sup to end sup procedure
            Else
            .Recordset.MoveNext 'move next record or file
            End If
            End With
       Loop
    End If
        MsgBox "Invalid Password, try again!", , "Login"
        txtpassword.Text = ""
        txtpassword.SetFocus
     
    
End Sub

