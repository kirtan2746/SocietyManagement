VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form sign 
   Caption         =   "sign in"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   480
      TabIndex        =   10
      Top             =   360
      Width           =   8055
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "MailBox"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtsub 
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
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtfrom 
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
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
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
      Left            =   6840
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmdpre 
      Caption         =   "Previous"
      Enabled         =   0   'False
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
      Left            =   3600
      TabIndex        =   4
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
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
      Left            =   720
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton cmdcomp 
      Caption         =   "Compose"
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
      Left            =   720
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtbr 
      Height          =   3615
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393217
      MaxLength       =   500
      TextRTF         =   $"sign.frx":0000
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
   Begin VB.CommandButton view 
      Caption         =   "Inbox"
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
      Left            =   720
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label lblsub 
      Caption         =   "Subect :"
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblfrom 
      Caption         =   "From :"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "sign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mail As String
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ps As Integer
Dim lb As Integer

Private Sub cmdback_Click()
memlgn.Show
Unload Me
End Sub

Private Sub cmdcomp_Click()
send.Show
Unload Me
End Sub

Private Sub cmdnxt_Click()
If rs.EOF <> True And ps <> lb Then
rs.MoveNext

ps = ps + 1
rtbr.Text = rs!email
txtsub.Text = rs!subject
txtfrom.Text = rs!fromid
'MsgBox lb
If rs.EOF <> True And ps <> lb Then
'.Caption = "NOTICE BOARD        " + Str(rs!ndate)
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
rtbr.Text = rs!email
txtsub.Text = rs!subject
txtfrom.Text = rs!fromid
'MsgBox ps
End If

If rs.BOF <> True And ps <> 1 Then
'lbldt.Caption = "NOTICE BOARD        " + Str(rs!ndate)
cmdnxt.Enabled = True


Else
cmdnxt.Enabled = True
cmdpre.Enabled = False
End If
End Sub


Private Sub view_Click()


'sqlquery = "select email from mail where memid='" & mname & "'"
'rs.Open sqlquery, con, adOpenDynamic, adLockPessimistic

    'rs.MoveFirst
    'rtbr.Text = rs.Fields("email").Value
    If ps <> 0 Then
    lblfrom.Visible = True
    txtfrom.Visible = True
    lblsub.Visible = True
    txtsub.Visible = True
    rtbr.Text = rs!email
    txtfrom.Text = rs!fromid
    txtsub.Text = rs!subject
    Else
    rtbr.Locked = True
    MsgBox ("NO EMAILS")
    
    End If
End Sub

Private Sub Form_Load()

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
rs.Open "select * from mail where memid='" & mname & "'", con, adOpenDynamic, adLockPessimistic
ps = rs.RecordCount
lb = ps
'MsgBox ps
If rs.RecordCount > 1 Then
rs.MoveLast
cmdpre.Enabled = True
cmdnxt.Enabled = False
'rtb.Text = rs!notice

ElseIf rs.RecordCount = 1 Then
rs.MoveLast
cmdpre.Enabled = False
cmdnxt.Enabled = False
'rs.MoveFirst
'rtbr.Text = rs!notice
Else
rtbr.Text = ""
cmdpre.Enabled = False
cmdnxt.Enabled = False
End If


End Sub

