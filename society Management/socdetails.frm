VERSION 5.00
Begin VB.Form socdetails 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton home 
      Caption         =   "HOME"
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
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame details 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   2520
      TabIndex        =   0
      Top             =   1200
      Width           =   11655
      Begin VB.Label Label2 
         Caption         =   "Society Name :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "  Address :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "No. of buildings :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "No. of flats :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Type of Flats :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Project By :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Phone no :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Aundh Road, Khadki, Near St. Thomas Church, Pune-411020."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label12 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label13 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "1/2/3 BHK Terrace Flats"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Goyal Patel Group"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label Label17 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   3
         Top             =   4320
         Width           =   3735
      End
      Begin VB.Label Label18 
         Caption         =   "Authorities :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label19 
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   4920
         Width           =   4815
      End
   End
   Begin VB.Label socdetails 
      Caption         =   "       DETAILS OF SOCIETY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   17
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "socdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim con As ADODB.Connection
Dim s As String
Private Sub Form_Load()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
rs.Open "select count(bname) as count ,sum (nof) as total  from build", con, adOpenDynamic, adLockPessimistic
Label10.Caption = start.Label5.Caption
Label12.Caption = rs!Count
Label13.Caption = rs!total
rs.Close
rs.Open "member", con, adOpenDynamic, adLockPessimistic
Do While (Not rs.EOF)
    If rs!Id = 1 Then
    Exit Do
    End If
    rs.MoveNext
Loop
Label19.Caption = rs!fname & " " & rs!lname
If rs!phnno = "" Then
Label17.Caption = "-"
Else
Label17.Caption = Str(rs!phnno)
End If
rs.Close
End Sub

Private Sub home_Click()
memlgn.Show
Unload Me
End Sub

