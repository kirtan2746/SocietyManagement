VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form send 
   Caption         =   "mail"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   7095
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "mail.frx":0000
      Left            =   1920
      List            =   "mail.frx":0002
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   5175
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott"
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
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
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
      Left            =   240
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtto 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox Txtsub 
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   3240
      Width           =   5175
   End
   Begin RichTextLib.RichTextBox rtbmail 
      Height          =   3375
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      MaxLength       =   500
      TextRTF         =   $"mail.frx":0004
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
   Begin VB.Label Label3 
      Caption         =   "To :"
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
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Subject:"
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
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Compose:"
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
      Left            =   480
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
End
Attribute VB_Name = "send"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim cnt As Integer
Dim tosend As String
Dim subject As String
Dim eml As String
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub cmdback_Click()
sign.Show
Unload Me
End Sub

Private Sub cmdsend_Click()
If (txtto.Text = "" Or txtsub.Text = "" Or rtbmail.Text = "") Then

    MsgBox "Missing Fields"
Else
fromid = mname
tosend = txtto.Text
subject = txtsub.Text
eml = rtbmail.Text
con.Execute ("insert into mail values('" & tosend & "','" & fromid & "','" & subject & "','" & eml & "')")
MsgBox ("Successfully sent")
End If
End Sub


Private Sub Form_Load()

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
rs.Open "mail", con, adOpenDynamic, adLockPessimistic
cnt = rs.RecordCount
'rtbmail.Locked = True



End Sub

Private Sub List1_Click()
If List1.Text <> "" Then
txtto.Text = List1.Text
List1.Visible = False
'MsgBox List1.TabIndex
End If


End Sub

Private Sub txtto_Change()
List1.Clear
List1.Height = 0
letter = txtto.Text
Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger;"
    rs.Open "select * from member where memid like '" & letter & "%'", con, adOpenDynamic, adLockPessimistic


If txtto.Text <> "" Then
List1.Visible = True

Do While rs.EOF <> True
List1.AddItem (rs!memid)

If List1.Height <= 840 Then
List1.Height = List1.Height + 210

End If
send.Refresh
rs.MoveNext
Loop
Else
List1.Visible = False

End If
End Sub

