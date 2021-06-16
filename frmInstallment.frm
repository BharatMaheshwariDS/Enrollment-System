VERSION 5.00
Begin VB.Form frmInstallment 
   BackColor       =   &H00E0E0E0&
   Caption         =   "installment"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000B&
      Caption         =   "BACK"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Data dtInstallment 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "G:\FINAL PROGRAM\prepayment2000.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Installment Fees"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      Begin VB.CommandButton Command2 
         BackColor       =   &H8000000B&
         Caption         =   "PAY"
         Height          =   495
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000B&
         Caption         =   "Compute"
         Height          =   495
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Payment"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Balance"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Total Fees"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSTALLMENT"
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
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmInstallment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label3 = (Label1 + Text1)
End Sub

Private Sub Command3_Click()
frmInstallment.Hide
frmenrollment.Show
End Sub
