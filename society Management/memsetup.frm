VERSION 5.00
Begin VB.Form memsetup 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton memsetbk 
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
      Left            =   1800
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdcpas 
      Caption         =   "OK"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtcpas 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtnpas 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtopas 
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
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Conform Password :"
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
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "New Password :"
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
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Old Password :"
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
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Privacy Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "memsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Recordset
Dim con As Connection

Private Sub cmdcpas_Click()
Do While Not rs.EOF
    If rs!memid = mname And rs!Password = txtopas.Text Then
        Exit Do
    Else
    rs.MoveNext
    End If
Loop
    
If rs.EOF Then
    MsgBox ("INVALID PASSWORD!!!!!!")
ElseIf txtnpas.Text <> txtcpas.Text Then
    MsgBox ("Confimation Password Mismatch!!")
Else
    con.Execute ("update member set password='" & txtnpas.Text & "'where memid = '" & mname & "'")
    MsgBox ("Password changed successfully")
    rs.MoveNext
    Unload Me
End If
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
    rs.Open "member", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub memsetbk_Click()
Unload Me
End Sub
