VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form LoginRegist 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frDatas 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2055
      Index           =   0
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
      Begin VB.TextBox txtpassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtusername 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   105
         Width           =   1815
      End
      Begin VB.ComboBox cmbgroup 
         Height          =   315
         ItemData        =   "LoginRegist.frx":0000
         Left            =   1200
         List            =   "LoginRegist.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1545
         Width           =   1815
      End
      Begin VB.TextBox txtConfirm 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1065
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   570
         TabIndex        =   6
         Top             =   1545
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm"
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
         Left            =   450
         TabIndex        =   5
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UserName"
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
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   105
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2280
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton Command2 
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
      Left            =   5640
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   "Add User"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "LoginRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim msg As Integer
Adodc1.Refresh
msg = MsgBox("Save Record?", vbYesNo, "Messge")
If msg = vbYes Then
    
    Do Until Adodc1.Recordset.EOF
        'search if username already stored para walang duplication
        If txtusername.Text = Adodc1.Recordset.Fields("userid") Then
            MsgBox "Username Already Registered.", vbOKOnly, "Message"
            txtusername.Text = ""
            txtpassword.Text = ""
            txtConfirm.Text = ""
            txtusername.SetFocus
        Exit Sub
        Else
            Adodc1.Recordset.MoveNext
        End If
    Loop
    
    'if password and confirm is equal pwede nang magsave
    If txtpassword.Text <> txtConfirm.Text Or txtusername.Text = "" _
        Or txtpassword.Text = "" Or cmbgroup.Text = "" Then
        MsgBox "invalid", vbOKOnly, "Message"
    Else
        
        With Adodc1.Recordset
            .AddNew
            .Fields("userID") = txtusername.Text
            .Fields("password") = txtpassword.Text
            .Fields("usergroup") = cmbgroup.Text
            .Save
        End With
            
            frmUserGroup.Adodc1.Refresh
            Set frmUserGroup.DataGrid1.DataSource = frmUserGroup.Adodc1
            frmUserGroup.Adodc1.Refresh
            Unload Me
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
