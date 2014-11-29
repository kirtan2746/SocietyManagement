VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Member 
   Caption         =   "Member Details"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Member.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   9645
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6840
      Top             =   7320
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
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
      Connect         =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from member"
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
   Begin VB.CommandButton btn_display 
      Caption         =   "Display All Records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   16
      Top             =   8040
      Width           =   2895
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Member.frx":23534
      Height          =   5655
      Left            =   6840
      TabIndex        =   15
      Top             =   1680
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   9975
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "MEMID"
         Caption         =   "MEMID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FIRSTNAME"
         Caption         =   "FIRSTNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "MIDDLENAME"
         Caption         =   "MIDDLENAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LASTNAME"
         Caption         =   "LASTNAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "BNO"
         Caption         =   "BNO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FLATNO"
         Caption         =   "FLATNO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "OWNER"
         Caption         =   "OWNER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "TYPENAME"
         Caption         =   "TYPENAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "FTYPEID"
         Caption         =   "FTYPEID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6495
      Begin VB.TextBox member 
         DataField       =   "MEMID"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton btn_search 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   21
         Top             =   5160
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "TYPENAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   20
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         DataField       =   "OWNER"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Top             =   3600
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "BNO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton btn_delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton btn_update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton btn_add 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox flatno 
         DataField       =   "FLATNO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   11
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox lastname 
         DataField       =   "LASTNAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox middlename 
         DataField       =   "MIDDLENAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox firstname 
         DataField       =   "FIRSTNAME"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Flat Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Owner:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Flat No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Building No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Middle Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "  Member Details"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Menu menu_back 
      Caption         =   "&Back"
   End
End
Attribute VB_Name = "Member"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim memberid As Integer
Dim flatid As Integer

Private Sub btn_add_Click()
If (firstname.Text = "" Or lastname.Text = "" Or flatno.Text = "") Then
    MsgBox "Missing Fields"
Else
Dim isOwner As Integer
Dim flatnumber As Integer
If Check1.Value = 1 Then
    isOwner = 1
    Else
    isOwner = 0
End If
If Combo2.Text = "1 BHK" Then
flatid = 1
ElseIf Combo2.Text = "2 BHK" Then
flatid = 2
ElseIf Combo2.Text = "3 BHK" Then
flatid = 3
Else
MsgBox ("Missing Fields")
End If
Adodc1.RecordSource = "select memid from member where flatno=" & flatno.Text & " and bno=" & Combo1.Text & " "
Adodc1.Refresh
If member.Text = "" Then
flatnumber = CInt(flatno.Text)
con.Execute ("insert into member(firstname,middlename,lastname,bno,flatno,owner,typename,ftypeid) values('" & firstname.Text & "','" & middlename.Text & "','" & lastname.Text & "'," & Combo1.Text & "," & flatnumber & "," & isOwner & ",'" & Combo2.Text & "'," & flatid & ")")
MsgBox "Status Added Successfully"
Adodc1.Refresh
Else
MsgBox ("Entered Flat Number Already Allocated To Someone")
End If

'firstname.Text = ""
'middlename.Text = ""
'lastname.Text = ""
'flatno.Text = ""
'Combo1.Text = ""
'Combo2.Text = ""
'firstname.SetFocus
End If
End Sub

Private Sub btn_delete_Click()
firstname.Text = ""
middlename.Text = ""
lastname.Text = ""
Combo1.Text = ""
flatno.Text = ""
memberid = InputBox("Enter member ID whose data you want to delete...")
Adodc1.RecordSource = "select * from member where memid=" & memberid & ""
Adodc1.Refresh
If flatno.Text = "" Then
    Adodc1.Refresh
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    con.Execute ("delete from member where memid='" & memberid & "'")
    MsgBox "Record deleted Successfully"
    firstname.Text = ""
    middlename.Text = ""
    lastname.Text = ""
    Combo1.Text = ""
    flatno.Text = ""
    Combo2.Text = ""
    End If
End Sub

Private Sub btn_display_Click()
Adodc1.RecordSource = "select * from member"
Adodc1.Refresh
firstname.Text = ""
middlename.Text = ""
lastname.Text = ""
Combo1.Text = ""
flatno.Text = ""
Combo2.Text = ""
End Sub

Private Sub btn_search_Click()
firstname.Text = ""
middlename.Text = ""
lastname.Text = ""
Combo1.Text = ""
flatno.Text = ""
memberid = InputBox("Enter member ID whose data you want to delete...")
Adodc1.RecordSource = "select * from member where memid=" & memberid & ""
Adodc1.Refresh
If flatno.Text = "" Then
    Adodc1.Refresh
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
End If
End Sub

Private Sub btn_update_Click()
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Form_Load()
connectdb
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.ConnectionString = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
con.Open
rs.Open "Select * From member", con, adOpenDynamic
Adodc1.Caption = Adodc1.Recordset.RecordCount & " Rows"
Set rs = con.Execute("select bno from building")
While (Not rs.EOF)
   Combo1.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
Set rs = con.Execute("select typename from flattype")
While (Not rs.EOF)
   Combo2.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
    firstname.Text = ""
    middlename.Text = ""
    lastname.Text = ""
    Combo1.Text = ""
    flatno.Text = ""
    Combo2.Text = ""
End Sub

Private Sub menu_back_Click()
Unload Me
MDIForm1.Show
End Sub
