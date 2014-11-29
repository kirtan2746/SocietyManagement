VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form admin 
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   ScaleHeight     =   8790
   ScaleWidth      =   16020
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   375
      Left            =   14640
      Top             =   360
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from charges"
      Caption         =   "Adodc4"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   8400
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "tiger"
      RecordSource    =   "select * from worker"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton cmdset 
      Caption         =   "SETUP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8400
      Top             =   480
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
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
      Connect         =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
      OLEDBString     =   "Provider=OraOLEDB.Oracle.1;Password=tiger;Persist Security Info=True;User ID=scott"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "scott"
      Password        =   "tiger"
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
   Begin TabDlg.SSTab tab1 
      Height          =   7905
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   13944
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "NOTICES"
      TabPicture(0)   =   "administrator.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbldt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "rtb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdnxt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdpre"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdsav"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdnewntc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdrmv"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmded"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "MEMBERS"
      TabPicture(1)   =   "administrator.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "datag"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "WORKERS"
      TabPicture(2)   =   "administrator.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGrid1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "COMPLAINTS"
      TabPicture(3)   =   "administrator.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblfrom"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblsub"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label17"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "listsub"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtfrom"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "rtbcomp"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "BILLING"
      TabPicture(4)   =   "administrator.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "DataGrid2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmddisp"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      Begin VB.CommandButton cmddisp 
         Caption         =   "DISPLAY"
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
         Left            =   9960
         TabIndex        =   86
         Top             =   6720
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "administrator.frx":008C
         Height          =   5895
         Left            =   7320
         TabIndex        =   83
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "MEMID"
            Caption         =   "MEMBER ID"
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
            DataField       =   "FNAME"
            Caption         =   "NAME"
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
            DataField       =   "MONTH"
            Caption         =   "MONTH"
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
            DataField       =   "FTYPE"
            Caption         =   "FLAT TYPE"
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
            DataField       =   "CHARGEPAID"
            Caption         =   "CHARGES"
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
            DataField       =   "ISPAID"
            Caption         =   "STATUS"
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
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame5 
         Caption         =   "Society Charge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   360
         TabIndex        =   69
         Top             =   720
         Width           =   6855
         Begin VB.CommandButton cmdel 
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
            Left            =   4200
            TabIndex        =   85
            Top             =   5400
            Width           =   1695
         End
         Begin VB.CommandButton cmsearch 
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
            Left            =   2160
            TabIndex        =   84
            Top             =   5400
            Width           =   1575
         End
         Begin VB.ComboBox Combo6 
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
            Left            =   2760
            TabIndex        =   76
            Top             =   1440
            Width           =   2775
         End
         Begin VB.CommandButton cmadd 
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
            TabIndex        =   75
            Top             =   5400
            Width           =   1335
         End
         Begin VB.ComboBox Combo7 
            DataField       =   "MEMID"
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
            Left            =   2760
            TabIndex        =   74
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtftype 
            DataField       =   "TYPENAME"
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
            Left            =   2760
            TabIndex        =   73
            Top             =   2880
            Width           =   2775
         End
         Begin VB.TextBox totalcharge 
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
            Left            =   2760
            MaxLength       =   6
            TabIndex        =   72
            Top             =   3600
            Width           =   2775
         End
         Begin VB.TextBox txfname 
            DataField       =   "FIRSTNAME"
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
            Left            =   2760
            TabIndex        =   71
            Top             =   2160
            Width           =   2775
         End
         Begin VB.CheckBox Check1 
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
            Left            =   3000
            TabIndex        =   70
            Top             =   4320
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.Label Label24 
            Caption         =   "Member ID:"
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
            Left            =   480
            TabIndex        =   82
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Month:"
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
            Left            =   480
            TabIndex        =   81
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label22 
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
            TabIndex        =   80
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Total Charge:"
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
            TabIndex        =   79
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Name:"
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
            Left            =   480
            TabIndex        =   78
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Is Paid:"
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
            TabIndex        =   77
            Top             =   4440
            Width           =   1695
         End
      End
      Begin RichTextLib.RichTextBox rtbcomp 
         Height          =   3735
         Left            =   -65400
         TabIndex        =   66
         Top             =   2760
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   6588
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   500
         TextRTF         =   $"administrator.frx":00A1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtfrom 
         Height          =   375
         Left            =   -65400
         TabIndex        =   63
         Top             =   1920
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.ListBox listsub 
         Height          =   255
         ItemData        =   "administrator.frx":011F
         Left            =   -73680
         List            =   "administrator.frx":0121
         TabIndex        =   62
         Top             =   1920
         Width           =   5295
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "administrator.frx":0123
         Height          =   3375
         Left            =   -74880
         TabIndex        =   53
         Top             =   4080
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   5953
         _Version        =   393216
         Enabled         =   0   'False
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "WID"
            Caption         =   "Worker ID"
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
            DataField       =   "WNAME"
            Caption         =   "Worker Name"
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
            DataField       =   "ADDRESS"
            Caption         =   "Address"
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
            DataField       =   "SALARY"
            Caption         =   "Salary"
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
            DataField       =   "WTYPE"
            Caption         =   "Work Title"
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
            DataField       =   "AGE"
            Caption         =   "Age"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   6300.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2055.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   870.236
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Worker Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   15255
         Begin VB.CommandButton cmdcncl 
            Caption         =   "CANCEL"
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
            Left            =   12720
            TabIndex        =   87
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdup 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11280
            TabIndex        =   61
            Top             =   2760
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
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
            Left            =   10200
            TabIndex        =   46
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox doj 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "dd-mmm-yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   10200
            MaxLength       =   10
            TabIndex        =   45
            Top             =   1320
            Width           =   2655
         End
         Begin VB.ComboBox Combo4 
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
            Left            =   10200
            TabIndex        =   44
            Top             =   480
            Width           =   2655
         End
         Begin VB.CommandButton cmdsearch 
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
            Left            =   9600
            TabIndex        =   43
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmddel 
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
            Left            =   7680
            TabIndex        =   42
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdupdate 
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
            Left            =   5760
            TabIndex        =   41
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdadd 
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
            Left            =   3960
            TabIndex        =   40
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox sal 
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
            Left            =   3960
            MaxLength       =   5
            TabIndex        =   39
            Top             =   2160
            Width           =   2655
         End
         Begin VB.TextBox address 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   38
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox wname 
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
            Left            =   3960
            MaxLength       =   20
            TabIndex        =   37
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label18 
            Caption         =   "DD/MM/YYYY"
            Height          =   255
            Left            =   10320
            TabIndex        =   68
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Age:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   52
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Date Of Joining:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   51
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Worker Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8160
            TabIndex        =   50
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Salary:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   49
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   48
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Worker Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   47
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Member Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   15255
         Begin VB.TextBox txtphnno 
            Height          =   375
            Left            =   12600
            MaxLength       =   10
            TabIndex        =   59
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   10800
            TabIndex        =   54
            Top             =   1320
            Width           =   4095
            Begin VB.OptionButton Option2 
               Height          =   375
               Index           =   0
               Left            =   1320
               TabIndex        =   56
               Top             =   180
               Width           =   255
            End
            Begin VB.OptionButton Option2 
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   55
               Top             =   240
               Width           =   255
            End
            Begin VB.Label Label15 
               Caption         =   "Member"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2280
               TabIndex        =   58
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label14 
               Caption         =   "Admin"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame2 
            Height          =   615
            Left            =   10800
            TabIndex        =   31
            Top             =   360
            Width           =   4095
            Begin VB.OptionButton Option1 
               Height          =   375
               Index           =   1
               Left            =   3240
               TabIndex        =   35
               Top             =   120
               Width           =   255
            End
            Begin VB.OptionButton Option1 
               Height          =   375
               Index           =   0
               Left            =   1320
               TabIndex        =   32
               Top             =   120
               Width           =   255
            End
            Begin VB.Label lblown 
               Caption         =   "Owner"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   34
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label1 
               Caption         =   "Tenant"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2280
               TabIndex        =   33
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox txtfname 
            DataField       =   "FIRSTNAME"
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
            MaxLength       =   15
            TabIndex        =   23
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtmname 
            DataField       =   "MIDDLENAME"
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
            MaxLength       =   15
            TabIndex        =   22
            Top             =   1320
            Width           =   2055
         End
         Begin VB.TextBox txtlname 
            DataField       =   "LASTNAME"
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
            MaxLength       =   15
            TabIndex        =   21
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtfno 
            DataField       =   "FLATNO"
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
            Left            =   7440
            MaxLength       =   2
            TabIndex        =   20
            Top             =   1320
            Width           =   2055
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
            Height          =   375
            Left            =   2040
            TabIndex        =   19
            Top             =   2640
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
            Height          =   375
            Left            =   8400
            TabIndex        =   18
            Top             =   2640
            Width           =   1095
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
            Height          =   375
            Left            =   4080
            TabIndex        =   17
            Top             =   2640
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            DataField       =   "BNO"
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
            Left            =   7440
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
         Begin VB.ComboBox Combo2 
            DataField       =   "TYPENAME"
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
            Left            =   7440
            TabIndex        =   15
            Top             =   2040
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
            Height          =   375
            Left            =   6240
            TabIndex        =   14
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Phno no."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11280
            TabIndex        =   60
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   720
            TabIndex        =   27
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Building Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   26
            Top             =   480
            Width           =   1695
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
            Height          =   255
            Left            =   6360
            TabIndex        =   25
            Top             =   1320
            Width           =   735
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
            Height          =   255
            Left            =   6120
            TabIndex        =   24
            Top             =   2040
            Width           =   975
         End
      End
      Begin VB.CommandButton cmded 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -64560
         TabIndex        =   9
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton cmdrmv 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -64560
         TabIndex        =   8
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdnewntc 
         Caption         =   "Create New"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -64560
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdsav 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -64560
         TabIndex        =   6
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdpre 
         Caption         =   "Previous"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73800
         TabIndex        =   5
         Top             =   6480
         Width           =   1455
      End
      Begin VB.CommandButton cmdnxt 
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68640
         TabIndex        =   4
         Top             =   6480
         Width           =   1455
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   4935
         Left            =   -73800
         TabIndex        =   10
         Top             =   1260
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8705
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   500
         TextRTF         =   $"administrator.frx":0138
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid datag 
         Bindings        =   "administrator.frx":01B6
         Height          =   3255
         Left            =   -74880
         TabIndex        =   30
         Top             =   3960
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
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
         ColumnCount     =   8
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
            DataField       =   "FNAME"
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
            DataField       =   "MNAME"
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
            DataField       =   "LNAME"
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
            DataField       =   "BNAME"
            Caption         =   "BUILDING NAME"
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
            DataField       =   "FNO"
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
            DataField       =   "FTYPE"
            Caption         =   "FLAT TYPE"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2220.094
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2220.094
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
               ColumnWidth     =   1349.858
            EndProperty
         EndProperty
      End
      Begin VB.Label Label17 
         Caption         =   "COMPLAINT :"
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
         Left            =   -65400
         TabIndex        =   67
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label lblsub 
         Caption         =   "SUBJECTS :"
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
         Left            =   -73680
         TabIndex        =   65
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblfrom 
         Caption         =   "FROM : "
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
         Left            =   -65400
         TabIndex        =   64
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbldt 
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73680
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdlog 
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame frame 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Label lbladmenu 
         Alignment       =   2  'Center
         Caption         =   "ADMIN MENU"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim up As Integer
Dim opt, opt2 As Integer
Dim cntr, cntr2 As Integer
Dim flag As Integer
Dim i As Integer
Dim cnt As Integer
Dim con As ADODB.Connection
Dim rs, rs1 As ADODB.Recordset
Dim cs As Integer
Dim rt As String
Dim SQL, membid As String
Dim memberid, workerid As String

Private Sub btn_add_Click()
If (txtfname.Text = "" Or txtlname.Text = "" Or txtmname.Text = "" Or txtfno.Text = "" Or Combo1.Text = "" Or Combo2.Text = "") Then
    MsgBox "Missing Fields"
Else

Dim isOwner As Integer
Dim flatnumber As Integer
Dim fl As Integer
fl = 0
rs.Open "member", con, adOpenDynamic, adLockPessimistic
        
        
        Do While (Not rs.EOF)
            If rs!fno = txtfno.Text And rs!bname = Combo1.Text Then
                If rs!memid <> memberid Then
                fl = 1
                End If
                rs.MoveNext
            Else
            rs.MoveNext
            End If
        Loop
        
        If up <> 1 Then
        rs.AddNew
        rs!Password = "password"
        Else
        rs.MoveFirst
                Do While (Not rs.EOF)
                      If rs!memid = memberid Then
                            Exit Do
                      Else
                            rs.MoveNext
                      End If
                Loop
      
        End If
        
        If fl = 0 Then
            If opt = 1 Then
            rs!Owner = "No"
            ElseIf opt = 0 Then
            rs!Owner = "Yes"
            End If
            If opt2 = 1 Then
            rs!Id = 0
            ElseIf opt2 = 0 Then
            rs!Id = 1
            End If
            opt = 3
            opt2 = 3
                
        rs!memid = txtfname.Text & Combo1.Text & txtfno.Text
        rs!fname = txtfname.Text
        rs!mname = txtmname.Text
        rs!lname = txtlname.Text
        rs!bname = Combo1.Text
        rs!phnno = txtphnno.Text
        rs!fno = txtfno.Text
        rs!ftype = Combo2.Text
        If up <> 1 Then
        cntr = cntr + 1
        up = 0
        End If
        rs.MoveNext
       rs.Close
        
        MsgBox "Status Added Successfully"
        Adodc1.Refresh
        
        If cntr = 0 Then
        btn_delete.Enabled = False
        btn_search.Enabled = False
        btn_update.Enabled = False
        Else
        btn_delete.Enabled = True
        btn_search.Enabled = True
        btn_update.Enabled = True
        End If
        
        Else
            MsgBox ("Entered Flat Number Already Allocated To Someone")
        End If
txtfname.Text = ""
txtmname.Text = ""
txtlname.Text = ""
txtfno.Text = ""
Combo1.Text = ""
Combo2.Text = ""
txtphnno.Text = ""
Option1(0).Value = False
Option1(1).Value = False
Option2(0).Value = False
Option2(1).Value = False
'rs.MoveNext
'rs.Close
End If
End Sub

Private Sub btn_delete_Click()
    memberid = InputBox("Enter member ID whose data you want to delete...")
    Dim fl As Integer
    fl = 0
    rs.Open "member", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!memid = memberid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
        
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    rs.Delete
    MsgBox "Record deleted Successfully"
    cntr = cntr - 1
    End If
    If cntr = 0 Then
        btn_delete.Enabled = False
        btn_search.Enabled = False
        btn_update.Enabled = False
        Else
        btn_delete.Enabled = True
        btn_search.Enabled = True
        btn_update.Enabled = True
    End If
    Adodc1.Refresh
    admin.Refresh
    rs.Close
End Sub

Private Sub btn_search_Click()
If btn_search.Caption = "SEARCH" Then
memberid = InputBox("Enter member ID whose data you want to Search...")

Dim fl As Integer
    fl = 0
    rs.Open "member", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!memid = memberid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    txtfname.Text = rs!fname
    txtmname.Text = rs!mname
    txtlname.Text = rs!lname
    txtphnno.Text = rs!phnno
    Combo1.Text = rs!bname
    txtfno.Text = rs!fno
    Combo2.Text = rs!ftype
    rs.Close
    End If
    
        txtfname.Locked = True
        txtmname.Locked = True
        txtlname.Locked = True
        txtphnno.Locked = True
        txtfno.Locked = True
        Combo1.Locked = True
        Combo2.Locked = True
        btn_add.Enabled = False
        btn_update.Enabled = False
        btn_delete.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Option2(0).Enabled = False
        Option2(1).Enabled = False
        btn_search.Caption = "OK"
    Else
        txtfname.Locked = False
        txtmname.Locked = False
        txtlname.Locked = False
        txtphnno.Locked = False
        txtfno.Locked = False
        Combo1.Locked = False
        Combo2.Locked = False
        btn_add.Enabled = True
        btn_update.Enabled = True
        btn_delete.Enabled = True
        Option1(0).Enabled = True
        Option1(1).Enabled = True
        Option2(0).Enabled = True
        Option2(1).Enabled = True
        txtfname.Text = ""
        txtmname.Text = ""
        txtlname.Text = ""
        txtfno.Text = ""
        Combo1.Text = ""
        Combo2.Text = ""
        txtphnno.Text = ""
        btn_search.Caption = "SEARCH"
    End If
    admin.Refresh
End Sub

Private Sub btn_update_Click()
memberid = InputBox("Enter member ID whose data you want to update...")
Dim fl As Integer
    fl = 0
    rs.Open "member", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!memid = memberid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    txtfname.Text = rs!fname
    txtmname.Text = rs!mname
    txtlname.Text = rs!lname
    Combo1.Text = rs!bname
    txtfno.Text = rs!fno
    Combo2.Text = rs!ftype
    txtphnno.Text = rs!phnno
    up = 1
    btn_update.Enabled = False
    btn_delete.Enabled = False
    btn_search.Enabled = False
    End If
    rs.Close
    
End Sub

Private Sub cmadd_Click()
If (Combo7.Text = "" Or Combo6.Text = "" Or txfname.Text = "" Or txtftype.Text = "" Or totalcharge.Text = "") Then
    MsgBox "Missing Fields"
Else
If Check1.Value = 1 Then
    ispaid = "PAID"
    Else
    ispaid = "PENDING"
End If
Adodc4.RecordSource = "select * from charges "
'where month='" & Combo1.Text & "' and memid='" & Combo2.Text & " "
Adodc4.Refresh
'If member.Text = "" And monthofpay.Text = "" Then
con.Execute ("insert into charges(memid,fname,chargepaid,month,ftype,ispaid) values('" & Combo7.Text & "','" & txfname.Text & "'," & totalcharge.Text & ",'" & Combo6.Text & "','" & txtftype & "','" & ispaid & "')")
con.Execute ("insert into temp_charge(memid,fname,chargepaid,month,ftype,ispaid) values('" & Combo7.Text & "','" & txfname.Text & "'," & totalcharge.Text & ",'" & Combo6.Text & "','" & txtftype & "','" & ispaid & "')")
MsgBox "Status Added Successfully"
Adodc4.Refresh
'Else
'MsgBox ("You Have Already Paid Society Charge Of This Month")
'End If
Combo7.Text = ""
Combo6.Text = ""
txfname.Text = ""
txtftype.Text = ""
totalcharge.Text = ""
End If
End Sub

'//////////////////////////////////
Private Sub cmdadd_Click()
If (wname.Text = "" Or address.Text = "" Or sal.Text = "" Or doj.Text = "" Or Combo3.Text = "" Or Combo4.Text = "") Then
    MsgBox "Missing Fields"
Else
con.Execute ("insert into worker(wname,address,salary,wtype,doj,age) values('" & wname.Text & "','" & address.Text & "'," & sal.Text & ",'" & Combo4.Text & "','" & doj.Text & "'," & Combo3.Text & ")")
cntr2 = cntr2 + 1
End If
If cntr2 = 0 Then
        cmddel.Enabled = False
        cmdsearch.Enabled = False
        cmdupdate.Enabled = False
        cmdup.Enabled = False
Else
        cmddel.Enabled = True
        cmdsearch.Enabled = True
        cmdupdate.Enabled = True
        cmdup.Enabled = True
End If
Adodc2.Refresh

wname.Text = ""
address.Text = ""
sal.Text = ""
Combo3.Text = ""
Combo4.Text = ""
doj.Text = ""
End Sub

Private Sub cmdcncl_Click()
    wname.Text = ""
    address.Text = ""
    doj.Text = ""
    sal.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    cmdupdate.Enabled = True
    cmddel.Enabled = True
    cmdsearch.Enabled = True
    cmdadd.Enabled = True
    cmdup.Value = False
    cmdcncl.Visible = False
End Sub

Private Sub cmddel_Click()
workerid = InputBox("Enter worker ID whose data you want to delete...")
    Dim fl As Integer
    fl = 0
    rs.Open "worker", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!wid = workerid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
        
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    rs.Delete
    MsgBox "Record deleted Successfully"
    cntr2 = cntr2 - 1
    End If
    
    If cntr2 = 0 Then
        cmddel.Enabled = False
        cmdsearch.Enabled = False
        cmdupdate.Enabled = False
        cmdup.Enabled = False
    Else
        cmddel.Enabled = True
        cmdsearch.Enabled = True
        cmdupdate.Enabled = True
        cmdup.Enabled = True
    End If
    wname.Text = ""
    sal.Text = ""
    address.Text = ""
    Combo4.Text = ""
    Combo3.Text = ""
    doj.Text = ""
    
    
    Adodc2.Refresh
    rs.Close

End Sub

Private Sub cmddisp_Click()
Adodc4.RecordSource = "select * from charges"
Adodc4.Refresh
End Sub

Private Sub cmded_Click()
cmdsav.Visible = True
rtb.Locked = False
flag = 1
End Sub

Private Sub cmdel_Click()
Combo7.Text = ""
Combo6.Text = ""
txfname.Text = ""
txtftype.Text = ""
totalcharge.Text = ""
s = InputBox("Enter Member ID whose data you want to delete...")
Set rs = New ADODB.Recordset
 rs.Open "select * from charges where memid='" & s & "'", con, adOpenDynamic, adLockPessimistic
    If Not rs.EOF Then
      con.Execute ("delete from charges where memid='" & s & "'")
      
    Adodc4.Refresh
    MsgBox "Record deleted Successfully"
    Combo7.Text = ""
    Combo6.Text = ""
    txfname.Text = ""
    txtftype.Text = ""
    totalcharge.Text = ""
    Else
    'Adodc4.Refresh
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
  
    End If
    rs.Close
End Sub

Private Sub cmdlog_Click()
    Unload Me
    start.Refresh
    start.Show
End Sub


Private Sub cmdnewntc_Click()
cmdsav.Visible = True
lbldt.Visible = False
cmdrmv.Enabled = False
cmded.Enabled = False
rtb.Locked = False
rtb.Text = " "
rtb.SetFocus
flag = 0
End Sub

Private Sub cmdnxt_Click()
    cs = cs + 1
    rs1.MoveNext
    rtb.Text = rs1!notice
    lbldt.Caption = rs1!ndate
    cmdpre.Enabled = True
     If cs = cnt Then
        cmdnxt.Enabled = False
    End If
End Sub

Private Sub cmdpre_Click()
    
    cs = cs - 1
    rs1.MovePrevious
    rtb.Text = rs1!notice
    lbldt.Caption = rs1!ndate
    cmdnxt.Enabled = True
     If cs = 1 Then
        cmdpre.Enabled = False
    End If
    
End Sub

Private Sub cmdrmv_Click()
If cnt <> 0 Then
    If cs = cnt Then
        rs1.Delete
        rs1.MovePrevious
        cnt = cnt - 1
        cs = cs - 1
    Else
        rs1.Delete
        rs1.MoveNext
        cnt = cnt - 1
        If cs = cnt Then
            cmdnxt.Enabled = False
        End If
    End If
    If cnt <> 0 Then
        rtb.Text = rs1!notice
        lbldt.Caption = rs1!ndate
    Else
        rtb.Text = " "
    End If
    MsgBox ("Successfully Deleted")
    If cnt < 2 Then
        cmdpre.Enabled = False
        cmdnxt.Enabled = False
    End If
Else
    MsgBox ("No Record")
End If
flag = 0
End Sub

Private Sub cmdsav_Click()
If cnt <> 0 And flag = 0 Then
rs1.MoveLast
cmdpre.Enabled = True
End If
If flag = 0 Then
cnt = cnt + 1
rs1.AddNew
cs = cnt
cmdnxt.Enabled = False
End If
rs1!notice = rtb.Text
rs1!ndate = Date
MsgBox ("Done!!!")
cmdsav.Visible = False
lbldt.Visible = True
lbldt.Caption = rs1!ndate
cmdrmv.Enabled = True
cmded.Enabled = True
rtb.Locked = True
'If flag = 1 Then
rs1.Update
'End If
End Sub

Private Sub cmdsearch_Click()
If cmdsearch.Caption = "SEARCH" Then
workerid = InputBox("Enter worker ID whose data you want to Search...")

Dim fl As Integer
    fl = 0
    rs.Open "worker", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!wid = workerid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    wname.Text = rs!wname
    address.Text = rs!address
    doj.Text = rs!doj
    sal.Text = rs!salary
    Combo3.Text = rs!age
    Combo4.Text = rs!wtype
    rs.Close
    End If
    
        wname.Locked = True
        address.Locked = True
        doj.Locked = True
        sal.Text = True
        Combo3.Locked = True
        Combo4.Locked = True
        cmdadd.Enabled = False
        cmdupdate.Enabled = False
        cmddel.Enabled = False
        cmdsearch.Caption = "OK"
    Else
        wname.Locked = False
        address.Locked = False
        doj.Locked = False
        sal.Text = False
        Combo3.Locked = False
        Combo4.Locked = False
        cmdadd.Enabled = True
        cmdupdate.Enabled = True
        cmddel.Enabled = True
        wname.Text = ""
        address.Text = ""
        doj.Text = ""
        sal.Text = ""
        Combo3.Text = ""
        Combo4.Text = ""
        cmdsearch.Caption = "SEARCH"
    End If
    admin.Refresh
End Sub

Private Sub cmdset_Click()
    setup.tab2.Tab = 0
    setup.Show
End Sub

Private Sub cmdup_Click()
rs.Open "worker", con, adOpenDynamic, adLockPessimistic
    Do While Not rs.EOF
            If rs!wid = workerid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
    rs!wname = wname.Text
    rs!address = address.Text
    rs!doj = doj.Text
    rs!salary = sal.Text
    rs!age = Combo3.Text
    rs!wtype = Combo4.Text
    
    wname.Text = ""
    address.Text = ""
    doj.Text = ""
    sal.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    rs.MoveNext
    Adodc2.Refresh
    rs.Close
    cmdupdate.Enabled = True
    cmddel.Enabled = True
    cmdsearch.Enabled = True
    cmdadd.Enabled = True
    cmdup.Value = False
    cmdcncl.Visible = False
End Sub

'////////////////////////////////////////////////////////////////////////////
Private Sub cmdupdate_Click()
workerid = InputBox("Enter member ID whose data you want to update...")
Dim fl As Integer
    fl = 0
    rs.Open "worker", con, adOpenDynamic, adLockPessimistic
     Do While Not rs.EOF
    
            If rs!wid = workerid Then
                fl = 1
                Exit Do
            Else
            rs.MoveNext
            End If
    Loop
    If fl = 0 Then
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
    Else
    wname.Text = rs!wname
    address.Text = rs!address
    doj.Text = rs!doj
    sal.Text = rs!salary
    Combo3.Text = rs!age
    Combo4.Text = rs!wtype
    End If
    rs.Close
cmdup.Visible = True
cmdcncl.Visible = True
cmdadd.Enabled = False
cmdsearch.Enabled = False
cmddel.Enabled = False
cmddel.Enabled = False
End Sub




Private Sub cmsearch_Click()
Adodc4.Refresh
membid = InputBox("Enter member ID to be searched...")
Adodc4.RecordSource = "select * from charges where memid='" & membid & "'"
Adodc4.Refresh
Set rs = New ADODB.Recordset
 rs.Open "select * from charges where memid='" & membid & "'", con, adOpenDynamic, adLockPessimistic
    If rs!memid = Null Then
    Adodc4.Refresh
    MsgBox " Record Not Found. ", vbExclamation, " Not Found "
End If
rs.Close
End Sub
Private Sub Combo7_Click()
 Set rs = New ADODB.Recordset
 rs.Open "select fname,ftype from member where memid ='" & Combo7.Text & "'", con, adOpenDynamic, adLockPessimistic
 txfname = rs!fname
 txtftype = rs!ftype
 rs.Close
End Sub

Private Sub Form_Load()
        Set con = New ADODB.Connection
        Set rs = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        rs1.Open "notices", con, adOpenDynamic, adLockPessimistic
        cmdsav.Visible = False
        cnt = rs1.RecordCount
        rtb.Locked = True
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
    rs1.MoveLast
    
    cs = cnt
    rtb.Text = rs1!notice
    lbldt.Caption = rs1!ndate
    
    End If
    flag = 0
    'rs.Close
End Sub

Private Sub listsub_Click()
    If listsub.Text <> "" Then
    lblfrom.Visible = True
    txtfrom.Visible = True
    rtbcomp.Visible = True
'    lbld.Visible = True
    
    'Set con = New ADODB.Connection
    'Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
     '   con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        rs.Open "select * from complaints where subject= '" & listsub.Text & "'", con, adOpenDynamic, adLockPessimistic
        
        txtfrom.Text = rs!memid
        rtbcomp.Text = rs!comp
'        lbld.Caption = rs!mdate
        l = 1
        rs!opened = l
        rs.MoveNext
        rs.Close
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
opt = Index
End Sub


Private Sub Option2_Click(Index As Integer)
opt2 = Index
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
 If tab1.Tab = 0 Then
        'Set con = New ADODB.Connection
        'Set rs = New ADODB.Recordset
        'rs.CursorLocation = adUseClient
        'con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        'rs.Open "notices", con, adOpenDynamic, adLockPessimistic
        cmdsav.Visible = False
        cnt = rs1.RecordCount
        rtb.Locked = True
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
    rs1.MoveLast
    
    cs = cnt
    rtb.Text = rs1!notice
    lbldt.Caption = rs1!ndate
    flag = 0
    End If
    'rs.Close
    '////////////////////////////////////////////////////////////////////////////////////////////////////
    ElseIf tab1.Tab = 1 Then
        'Set con = New ADODB.Connection
        'Set rs = New ADODB.Recordset
        'rs.CursorLocation = adUseClient
        'con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        Adodc1.Caption = Adodc1.Recordset.RecordCount & " Rows"
        rs.Open "Select * From member", con, adOpenDynamic

        Set rs = con.Execute("select bname from build")
        Combo1.Clear
        While (Not rs.EOF)
            Combo1.AddItem rs(0)
            rs.MoveNext
        Wend
        rs.Close
        Set rs = con.Execute("select ftype from flat")
        Combo2.Clear
        While (Not rs.EOF)
            Combo2.AddItem rs(0)
            rs.MoveNext
        Wend
        rs.Close
        opt = 3
        rs.CursorLocation = adUseClient
        rs.Open "member", con, adOpenDynamic, adLockPessimistic
        cntr = rs.RecordCount
        If cntr = 0 Then
        btn_delete.Enabled = False
        btn_search.Enabled = False
        btn_update.Enabled = False
        Else
        btn_delete.Enabled = True
        btn_search.Enabled = True
        btn_update.Enabled = True
        'rs.Close
        End If
        rs.Close
        '///////////////////////////////////////////////////////////////////////////////////////////////
        ElseIf tab1.Tab = 2 Then
        'Set con = New ADODB.Connection
        'Set rs = New ADODB.Recordset
        'rs.CursorLocation = adUseClient
        Adodc2.Caption = Adodc2.Recordset.RecordCount & " Rows"
        'con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        rs.Open "select * from worker", con, adOpenDynamic
               
        cntr2 = rs.RecordCount
        If cntr2 = 0 Then
        cmddel.Enabled = False
        cmdsearch.Enabled = False
        cmdupdate.Enabled = False
        cmdup.Enabled = False
        Else
        cmddel.Enabled = True
        cmdsearch.Enabled = True
        cmdupdate.Enabled = True
        cmdup.Enabled = True
        End If
        Combo4.Clear
        Combo4.AddItem ("Electrician")
        Combo4.AddItem ("Plumber")
        Combo4.AddItem ("Security")
        Combo4.AddItem ("Janitor")
        
        Combo3.Clear
        i = 18
        While i <= 60
            Combo3.AddItem (i)
            i = i + 1
        Wend
        rs.Close
        '//////////////////////////////////////////////////////////////////////////////////////////////
        ElseIf tab1.Tab = 3 Then
        'Set con = New ADODB.Connection
        'Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        'con.Open "Provider=OraOLEDB.Oracle.1;Persist Security Info=False;User ID=scott;Password=tiger"
        rs.Open "select * from complaints", con, adOpenDynamic, adLockPessimistic
        listsub.Clear
        listsub.Height = 210
        While Not rs.EOF
        listsub.AddItem (rs!subject)
        rs.MoveNext
        If listsub.Height <= 1050 Then
        listsub.Height = listsub.Height + 210
        End If
        Wend
        If rs.RecordCount <> 0 Then
        rs.MoveLast
        End If
        txtfrom.Text = ""
        rtbcomp.Text = ""
        
        rs.Close
          ElseIf tab1.Tab = 4 Then
          rs.CursorLocation = adUseClient
         rs.Open "select * from charges", con, adOpenDynamic, adLockPessimistic
        Combo6.Clear
        Combo6.AddItem "January"
        Combo6.AddItem "February"
        Combo6.AddItem "March"
        Combo6.AddItem "April"
        Combo6.AddItem "May"
        Combo6.AddItem "June"
        Combo6.AddItem "July"
        Combo6.AddItem "August"
        Combo6.AddItem "September"
        Combo6.AddItem "October"
        Combo6.AddItem "November"
        Combo6.AddItem "December"
        Set rs = con.Execute("select memid from member")
        Combo7.Clear
        While (Not rs.EOF)
           Combo7.AddItem rs(0)
            rs.MoveNext
        Wend
        rs.Close
        Combo7.Text = ""
        Combo6.Text = ""
        txfname.Text = ""
        txtftype.Text = ""
        totalcharge.Text = ""
        con.Execute ("delete from temp_charge")
    End If
End Sub
