VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form complaints 
   Caption         =   "Complaints"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   9495
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Complaint Box"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   6135
      End
   End
   Begin VB.CheckBox cb 
      Height          =   615
      Left            =   9360
      TabIndex        =   7
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
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
      Height          =   405
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1560
      Width           =   6375
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
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   6600
      Width           =   1050
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
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdsav 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdnewntc 
      Caption         =   "Create New"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtbcomp 
      Height          =   4095
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      _Version        =   393217
      ReadOnly        =   -1  'True
      MaxLength       =   500
      TextRTF         =   $"complaints.frx":0000
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
      Caption         =   "Complaint :"
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
      Left            =   495
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Subject :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "VIEWED"
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
      Left            =   9600
      TabIndex        =   8
      Top             =   2100
      Width           =   975
   End
End
Attribute VB_Name = "complaints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cnt As Integer
Dim cs As Integer

Private Sub cmdback_Click()
memlgn.Show
Unload Me
End Sub

Private Sub cmdnewntc_Click()
rtbcomp.Locked = False
txtsub.Locked = False
cmdsav.Enabled = True

rtbcomp.Text = " "
txtsub.Text = " "
txtsub.SetFocus
End Sub

Private Sub cmdnxt_Click()
    cs = cs + 1
    rs.MoveNext
    rtbcomp.Text = rs!comp
    txtsub.Text = rs!subject
    cmdpre.Enabled = True
     If cs = cnt Then
        cmdnxt.Enabled = False
    End If
    If rs!opened <> 0 Then
        cb.Enabled = True
        cb.Value = 1
        cb.Enabled = False
    Else
        cb.Enabled = True
        cb.Value = 0
        cb.Enabled = False
    End If
    
End Sub

Private Sub cmdpre_Click()
    cs = cs - 1
    rs.MovePrevious
    rtbcomp.Text = rs!comp
    txtsub.Text = rs!subject
    cmdnxt.Enabled = True
     If cs = 1 Then
        cmdpre.Enabled = False
    End If
    If rs!opened <> 0 Then
        cb.Enabled = True
        cb.Value = 1
        cb.Enabled = False
    Else
        cb.Enabled = True
        cb.Value = 0
        cb.Enabled = False
    End If
    
End Sub


Private Sub cmdsav_Click()
If (rtbcomp.Text = "" Or txtsub.Text = "") Then
MsgBox ("Enter all fields")
Else
If cnt <> 0 Then
rs.MoveLast
cmdpre.Enabled = True
End If
cnt = cnt + 1
cs = cnt
rs.AddNew
cmdnxt.Enabled = False

rs!comp = rtbcomp.Text
rs!mdate = Date
rs!memid = mname
rs!subject = txtsub.Text
cb.Enabled = True
cb.Value = 0
cb.Enabled = False
rs!opened = 0
MsgBox ("Complaint Sent!!!")
rtbcomp.Locked = True
txtsub.Locked = True
cmdsav.Enabled = False
'rs.MoveNext
rs.Update
End If
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        rs.Open "select * from complaints where memid ='" & mname & "' ", con, adOpenDynamic, adLockPessimistic
    
        cnt = rs.RecordCount
        rtbcomp.Locked = True
        txtsub.Locked = True
        cmdsav.Enabled = False
           If cnt = 0 Then
                cmdnxt.Enabled = False
                cmdpre.Enabled = False
                cs = 0
            
           ElseIf cnt = 1 Then
                cmdnxt.Enabled = False
                cmdpre.Enabled = False
            Else
                cmdpre.Enabled = True
                cmdnxt.Enabled = False
            End If
    If cnt <> 0 Then
    rs.MoveLast
    If rs!opened <> 0 Then
        cb.Enabled = True
        cb.Value = 1
        cb.Enabled = False
    Else
        cb.Enabled = True
        cb.Value = 0
        cb.Enabled = False
    End If
    
    cs = cnt
    rtbcomp.Text = rs!comp
    txtsub.Text = rs!subject
    End If
    
End Sub
