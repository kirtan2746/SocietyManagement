VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form memlgn 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12450
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   840
      TabIndex        =   12
      Top             =   120
      Width           =   10095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         TabIndex        =   13
         Top             =   480
         Width           =   6975
      End
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdwd 
      Caption         =   "Worker Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdmail 
      Caption         =   "Mail Box"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   6720
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   6240
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
      CommandType     =   8
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Data Source=melvin"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Data Source=melvin"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton cmdno 
      Caption         =   "Directory"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmddtls 
      Caption         =   "Society Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2640
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdcmplt 
      Caption         =   "Complaints"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdnxt 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   2
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdpre 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   3495
      Left            =   6840
      TabIndex        =   0
      Top             =   2640
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"memlgn.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbldt 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label lblwel 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "memlgn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ps As Integer
Dim hb As Integer
Dim lb As Integer

Private Sub cmdback_Click()
Unload Me
start.Refresh
start.Show

End Sub

Private Sub cmdcmplt_Click()
complaints.Show
Unload Me
End Sub

Private Sub cmddtls_Click()
socdetails.Show
Unload Me
End Sub
Private Sub cmdmail_Click()
sign.Show
Unload Me
End Sub

Private Sub cmdno_Click()
direc.Show
Unload Me
End Sub

Private Sub cmdnxt_Click()
If rs.EOF <> True And ps <> lb Then
rs.MoveNext
'MsgBox ps
ps = ps + 1
rtb1.Text = rs!notice
lbldt.Caption = "NOTICE BOARD        " + Str(rs!ndate)
If rs.EOF <> True And ps <> lb Then

cmdpre.Enabled = True
Else
cmdpre.Enabled = True
cmdnxt.Enabled = False
End If
End If
End Sub

Private Sub cmdpre_Click()
If rs.BOF <> True And ps <> 1 Then
rs.MovePrevious
ps = ps - 1
rtb1.Text = rs!notice
lbldt.Caption = "NOTICE BOARD        " + Str(rs!ndate)
'MsgBox ps
End If
If rs.BOF <> True And ps <> 1 Then
cmdnxt.Enabled = True
Else
cmdnxt.Enabled = True
cmdpre.Enabled = False
End If
End Sub


Private Sub cmdset_Click()
memsetup.Show
End Sub

Private Sub cmdwd_Click()
Unload Me
wdtail.Show
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger;"
rs.Open "notices", con, adOpenDynamic, adLockPessimistic
ps = rs.RecordCount

'MsgBox ps
lb = ps
lblwel.Caption = "Welcome  " + mname
If rs.RecordCount > 1 Then
rs.MoveLast
cmdnxt.Enabled = False
rtb1.Text = rs!notice
lbldt.Caption = "NOTICE BOARD        " + Str(rs!ndate)

ElseIf rs.RecordCount = 1 Then
rs.MoveLast
cmdpre.Enabled = False
cmdnxt.Enabled = False
'rs.MoveFirst
rtb1.Text = rs!notice
lbldt.Caption = "NOTICE BOARD        " + Str(rs!ndate)
lblwel.Caption = "Welcome  " + mname
Else
rtb1.Text = ""
lbldt.Caption = " NO CURRENT NOTICE"
cmdpre.Enabled = False
cmdnxt.Enabled = False
End If
End Sub

