VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form setup 
   Caption         =   "Setup"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "tiger"
      RecordSource    =   "select * from build"
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
   Begin TabDlg.SSTab tab2 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "setup.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fr1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Flat Type"
      TabPicture(1)   =   "setup.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(4)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(5)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(6)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(7)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(8)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label2(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label2(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label2(2)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label2(3)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label2(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label2(5)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label2(6)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label2(7)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label2(8)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "cmdfadd"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Check1(0)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Check1(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Check1(2)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Check1(3)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Check1(4)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Check1(5)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Check1(6)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Check1(7)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Check1(8)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   34
         Top             =   5520
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   25
         Top             =   4920
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   24
         Top             =   4320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   23
         Top             =   3720
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   22
         Top             =   3120
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   21
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   20
         Top             =   1920
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   19
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdfadd 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   17
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame fr1 
         Caption         =   "Building Settings"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6975
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   6015
         Begin VB.CommandButton cmdcncle 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   16
            Top             =   6360
            Width           =   975
         End
         Begin VB.CommandButton cmdmod 
            Caption         =   "Modify"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   14
            Top             =   6360
            Width           =   975
         End
         Begin VB.CommandButton cmdmodok 
            Caption         =   "OK"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   13
            Top             =   6360
            Width           =   1095
         End
         Begin VB.TextBox txtnob 
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdnob 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame fbinfo 
            Caption         =   "Addtional Info."
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   240
            TabIndex        =   3
            Top             =   1080
            Width           =   5535
            Begin VB.TextBox txtnof 
               Height          =   285
               Left            =   3000
               MaxLength       =   2
               TabIndex        =   6
               Top             =   1080
               Width           =   735
            End
            Begin VB.TextBox txtbn 
               Height          =   285
               Left            =   3000
               MaxLength       =   3
               TabIndex        =   5
               Top             =   480
               Width           =   1335
            End
            Begin VB.CommandButton cmdsav 
               Caption         =   "Save"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               TabIndex        =   4
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label lblnof 
               Caption         =   "No. of Flats :"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1440
               TabIndex        =   8
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label lblbn 
               Caption         =   "Name of the building :"
               BeginProperty Font 
                  Name            =   "Georgia"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   7
               Top             =   480
               Width           =   2175
            End
            Begin VB.Label lblbno 
               Height          =   375
               Left            =   1560
               TabIndex        =   9
               Top             =   240
               Width           =   2535
            End
         End
         Begin MSDataGridLib.DataGrid Datag 
            Bindings        =   "setup.frx":0038
            Height          =   2535
            Left            =   240
            TabIndex        =   15
            Top             =   3600
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   4471
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   17
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
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Building Information"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "BNAME"
               Caption         =   "Name"
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
               DataField       =   "NOF"
               Caption         =   "No. of flats"
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
                  ColumnWidth     =   2729.764
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2459.906
               EndProperty
            EndProperty
         End
         Begin VB.Label lblnob 
            Caption         =   "No of building/s :"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Label Label2 
         Caption         =   "DUP"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   4920
         TabIndex        =   44
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "PENT"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   4920
         TabIndex        =   43
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "4 BHK T"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   4920
         TabIndex        =   42
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "3 BHK T"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   4920
         TabIndex        =   41
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "2 BHK T"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   4920
         TabIndex        =   40
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "4 BHK"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4920
         TabIndex        =   39
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "3 BHK"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4920
         TabIndex        =   38
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "2 BHK"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   4920
         TabIndex        =   37
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "1 BHK"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4920
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Duplex"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1320
         TabIndex        =   35
         Top             =   5640
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Pent House"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1320
         TabIndex        =   33
         Top             =   5040
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "4 - Bed Hall Kitchen Terrace"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1320
         TabIndex        =   32
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "3 - Bed Hall Kitchen Terrace"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   1320
         TabIndex        =   31
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "2 - Bed Hall Kitchen Terrace"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1320
         TabIndex        =   30
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "4 - Bed Hall Kitchen"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1320
         TabIndex        =   29
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "3 - Bed Hall Kitchen"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   28
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "2 - Bed Hall Kitchen"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   27
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "1 - Bed Hall Kitchen"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   26
         Top             =   840
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "FINISH"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   7800
      Width           =   1215
   End
End
Attribute VB_Name = "setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cntr As Integer
Dim tmp As Integer
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim i As Integer
Dim pos As Integer



Private Sub cmdcan_Click()
    setup.Hide
    Unload Me
End Sub

Private Sub cmdbok_Click()
rs.Open "flat", con, adOpenDynamic, adLockPessimistic
Do While Not rs.EOF
    If Combo1.Text = rs!ftype Then
        rs!scharge = txtscharge.Text
        rs!echarge = txtecharge.Text
        rs!wcharge = txtwcharge.Text
    End If
    rs.MoveNext
    Loop
End Sub

Private Sub cmdcncle_Click()
Adodc1.Refresh
Datag.AllowDelete = False
Datag.AllowUpdate = False
Datag.Enabled = False
End Sub

Private Sub cmdfadd_Click()
rs.Open "flat", con, adOpenDynamic, adLockPessimistic
i = 0
Do While (i < 9)
If Check1(i).Value = 1 Then
rs.AddNew
rs!ftype = Label2(i).Caption
Check1(i).Value = 0
Check1(i).Enabled = False
rs.Update
End If
i = i + 1
Loop
rs.Update
rs.Close
End Sub

Private Sub cmdmod_Click()
Datag.AllowDelete = True
Datag.AllowUpdate = True
cmdmodok.Enabled = True
Datag.Enabled = True
End Sub

Private Sub cmdmodok_Click()
Datag.AllowDelete = False
Datag.AllowUpdate = False
Datag.Enabled = False
End Sub

Private Sub cmdnob_Click()
If txtnob.Text <> "" And Val(txtnob.Text) > 0 Then
cntr = Val(txtnob.Text)
If cntr >= 30 Then
txtnob.Text = 30
End If
tmp = tmp + 1
fbinfo.Visible = True
lblbno.Caption = "Building No :" + Str(tmp)
txtnob.Locked = True
cmdsav.Enabled = False
Else
MsgBox ("Insert a value")
txtnob.SetFocus
End If
End Sub

Private Sub cmdok_Click()
Unload Me
End Sub


Private Sub cmdsav_Click()

rs.Open
rs.AddNew
rs!bname = txtbn.Text
rs!nof = txtnof.Text
rs.MoveNext
Adodc1.Refresh
txtbn.Text = ""
txtnof.Text = ""
cmdsav.Enabled = False
tmp = tmp + 1
If tmp <= cntr Then
    lblbno.Caption = "Building No :" + Str(tmp)
Else
    lblbno.Caption = "Done!!"
    txtbn.Enabled = False
    txtnof.Enabled = False
    cmdmod.Enabled = True
End If
rs.Close
End Sub




Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
    rs.Open "build", con, adOpenDynamic, adLockPessimistic
    fbinfo.Visible = False
    tmp = 0
    Datag.Enabled = False
    pos = rs.RecordCount
    If pos > 0 Then
    cmdmod.Enabled = True
    End If
    cmdnob.Enabled = False
rs.Close
End Sub




Private Sub tab2_Click(PreviousTab As Integer)
If tab2.Tab = 0 Then
rs.CursorLocation = adUseClient
    'con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
    rs.Open "build", con, adOpenDynamic, adLockPessimistic
    fbinfo.Visible = False
    tmp = 0
    Datag.Enabled = False
    pos = rs.RecordCount
    If pos > 0 Then
    cmdmod.Enabled = True
    End If
    cmdnob.Enabled = False
rs.Close
ElseIf tab2.Tab = 1 Then
rs.Open "flat", con, adOpenDynamic, adLockPessimistic
i = 0
If rs.RecordCount <> 0 Then
Do While (i < 9)

    Do While Not rs.EOF
        If rs!ftype = Label2(i).Caption Then
            Check1(i).Value = 0
            Check1(i).Enabled = False
        End If
        rs.MoveNext
    Loop
    rs.MoveFirst
i = i + 1
Loop
End If
rs.Close
End If
End Sub

Private Sub txtbn_Change()
If txtbn.Text <> "" And txtnof.Text <> "" Then
cmdsav.Enabled = True
End If
End Sub

Private Sub txtnob_Change()
cmdnob.Enabled = True
End Sub

Private Sub txtnof_Change()
If txtbn.Text <> "" And txtnof.Text <> "" Then
cmdsav.Enabled = True
End If
End Sub
 
