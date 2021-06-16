VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Cash 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   6600
      Width           =   1215
   End
   Begin VB.PictureBox PageReciept 
      BackColor       =   &H80000004&
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   12180
      TabIndex        =   1
      Top             =   720
      Width           =   12240
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         Left            =   9480
         TabIndex        =   31
         Text            =   "0"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   9480
         TabIndex        =   30
         Text            =   "0"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   25
         Left            =   9480
         TabIndex        =   28
         Text            =   "0"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   6960
         TabIndex        =   27
         Text            =   "Change"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         Left            =   6960
         TabIndex        =   26
         Text            =   "Cash"
         Top             =   2617
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   21
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Index           =   20
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   23
         Text            =   "Particulars"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   2760
         TabIndex        =   21
         Text            =   "Amount"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   6960
         TabIndex        =   20
         Text            =   "Balance"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   960
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   960
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   5040
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   3960
         TabIndex        =   16
         Text            =   "Payments"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   5280
         TabIndex        =   15
         Text            =   "Balance"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3960
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   2760
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5280
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   6960
         TabIndex        =   11
         Text            =   "Payment"
         Top             =   2126
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   6960
         TabIndex        =   10
         Text            =   "Remaining Balance"
         Top             =   3108
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   9480
         TabIndex        =   9
         Text            =   "0"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   9480
         TabIndex        =   8
         Text            =   "0"
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6720
         TabIndex        =   7
         Text            =   "Cashier:"
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2640
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label LblHeader 
         Alignment       =   2  'Center
         Caption         =   "Address"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   32
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label LblHeader 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   29
         Top             =   4440
         Width           =   10335
      End
      Begin VB.Line Line4 
         Index           =   9
         X1              =   6720
         X2              =   11280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         Index           =   8
         X1              =   6720
         X2              =   11280
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Line Line4 
         Index           =   7
         X1              =   6720
         X2              =   6720
         Y1              =   4080
         Y2              =   1440
      End
      Begin VB.Line Line4 
         Index           =   6
         X1              =   11280
         X2              =   11280
         Y1              =   1440
         Y2              =   4080
      End
      Begin VB.Line Line4 
         Index           =   5
         X1              =   7440
         X2              =   9600
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label LblHeader 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ENROLLMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   15.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   4260
         TabIndex        =   4
         Top             =   360
         Width           =   4005
      End
      Begin VB.Line Line4 
         Index           =   4
         X1              =   6960
         X2              =   10920
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line4 
         Index           =   3
         X1              =   3960
         X2              =   5160
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line4 
         Index           =   2
         X1              =   2760
         X2              =   3840
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line4 
         Index           =   1
         X1              =   5280
         X2              =   6480
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line4 
         Index           =   0
         X1              =   960
         X2              =   2640
         Y1              =   3720
         Y2              =   3720
      End
   End
   Begin MSAdodcLib.Adodc AdoPayment 
      Height          =   330
      Left            =   8520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tblpayments"
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
   Begin MSAdodcLib.Adodc AdoStudent 
      Height          =   330
      Left            =   11160
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form of payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2940
   End
End
Attribute VB_Name = "Cash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Name of temporary file

Const Text_File = "Text.TXT" 'costant text para sa

Private Sub cmdPrint_Click()
Dim msg As Integer

msg = MsgBox("Print Reciept?", vbYesNo, "Print")

If msg = vbYes Then
    'call printpage procedure
    PrintPage
End If

End Sub





'All printing calls start here

Function LoadTextToPrinter(Index As Integer)
    Dim NewLine As String
    
    On Error GoTo FileError
    
    ' Create a temporary file
    Open App.Path & "\" & Text_File For Output As #1
    Print #1, Text2(Index).Text
    ' close the output file
    Close #1
    
    ' Open the temporary file
    Open App.Path & "\" & Text_File For Input As #1
    
    ' position the text
    Printer.CurrentY = Text2(Index).Top
            
    ' print text one line at a time make the printers currentX = textbox's left position
    Do Until EOF(1)
                
        Line Input #1, NewLine
        Printer.CurrentX = Text2(Index).Left
        Printer.Print NewLine
                
    Loop
    ' close the input file
    Close #1
    
    ' delete the temporary file
    Kill App.Path & "\" & Text_File
    
    Exit Function
    
FileError:
    MsgBox "File Error!"
    
End Function

Function PrintPage()

    Dim i As Integer ' variable to call the textbox to print
    
    With Printer
        
        ' place heading on printer
        .Width = .Width
        .Height = PageReciept.Height
                
        'se all header array fonttype
        For i = 0 To 2 Step 1
        .Font = LblHeader(i).Font
        .FontSize = LblHeader(i).FontSize
        .FontBold = LblHeader(i).FontBold
        .FontUnderline = LblHeader(i).FontUnderline
        .CurrentY = LblHeader(i).Top
        .CurrentX = LblHeader(i).Left
        'print lblheater
        Printer.Print LblHeader(i).Caption
        Next
        
        For i = 0 To 9 Step 1
        Printer.Line (Line4(i).X1, Line4(i).Y1)-(Line4(i).X2, Line4(i).Y2)
        Next

        .Font = Text2(0).Font
        .FontSize = Text2(0).FontSize
        .FontBold = Text2(0).FontBold
        .FontUnderline = Text2(0).FontUnderline
            
       ' place all text files to the printer
        For i = 0 To 25 Step 1

            LoadTextToPrinter (i)

        Next
        
        .EndDoc
        
    End With
        
End Function

Private Sub cmdSave_Click()
Dim msg As Integer

msg = MsgBox("Save?", vbYesNo, "Message")

If msg = vbYes Then
'save payments and balance to student record
With AdoStudent.Recordset
    .Fields("balance") = Val(Text2(5).Text)
    .Fields("payments") = Val(Text2(6).Text) - .Fields("Payments")
    .Update
End With

MsgBox "Saved"
cmdPrint.Enabled = True
cmdSave.Enabled = False
Text2(3).Text = "Run Time: " & Time
End If

End Sub

Private Sub Form_Load()
Dim Amount As Integer
Dim stramt(1) As String

Amount = 0
'requiry adostudents where studentnumber = selected student number from enroll form or search form
AdoStudent.RecordSource = "select * from tblstudents where studentnumber = """ & stdno & """"
AdoStudent.Refresh
'where yearlevel = to yearlevel from formsearch or form enroll
AdoPayment.RecordSource = "select particulars, amount from tblpayments where yearlevel = """ & yrlvl & """"
AdoPayment.Refresh

With AdoStudent.Recordset
    'set text = form datafield might be
    'forname
    Text2(17).Text = .Fields("lastname") & ", " & .Fields("firstname") & " " & .Fields("middlename")
    'for student number
    Text2(16).Text = .Fields("studentnumber")
    'for yearlevel
    Text2(15).Text = .Fields("yearlevel")
    'for payments
    Text2(1).Text = .Fields("payments")
    'for balance
    Text2(10).Text = .Fields("Balance")
    'for balance
    Text2(7).Text = .Fields("Balance")
    'for payments
    Text2(12).Text = .Fields("Payments")
        
End With

AdoPayment.Refresh
'add all particular name and amount to multitext
Do Until AdoPayment.Recordset.EOF

    stramt(0) = AdoPayment.Recordset.Fields("Particulars")
    Text2(20).Text = Text2(20).Text & stramt(0) & Nl 'Nl call Newline Function
    
    stramt(1) = Str(AdoPayment.Recordset.Fields("Amount"))
    Text2(21).Text = Text2(21).Text & stramt(1) & Nl
    
    Amount = Amount + AdoPayment.Recordset.Fields("amount")
    AdoPayment.Recordset.MoveNext
    
Loop
'set total amount to text
Text2(11).Text = Amount

Text2(2).Text = "Run Date: " & Date
End Sub

Private Sub Text2_Change(Index As Integer)

'condition payment must be lessThan or equal to Balance
If Val(Text2(6).Text) > Val(Text2(7).Text) Then
    MsgBox "Payment Must Not Be Grater than Balance"
Else
    
    Text2(5).Text = Val(Text2(7).Text) - Val(Text2(6))
    LblHeader(2).Caption = "Amount Paid: " & cNumToWord(Text2(6).Text)
End If



End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    
    If Val(Text2(23).Text) >= Val(Text2(6).Text) Then
        'change = payment + cash
        Text2(25).Text = Val(Text2(23).Text) - Val(Text2(6).Text)
    
    End If
    
    If (Val(Text2(6).Text) <= Val(Text2(7).Text) And Val(Text2(23).Text) >= Val(Text2(6).Text)) Then
        
        cmdSave.Enabled = True
    
    Else
        
        cmdSave.Enabled = False
    
    End If
    
End If

End Sub


