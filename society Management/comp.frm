VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form comp 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listsub 
      Height          =   255
      ItemData        =   "comp.frx":0000
      Left            =   360
      List            =   "comp.frx":0002
      TabIndex        =   1
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox txtfrom 
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox rtbcomp 
      Height          =   3495
      Left            =   8640
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6165
      _Version        =   393217
      TextRTF         =   $"comp.frx":0004
   End
   Begin VB.Label lblfrom 
      Caption         =   "FROM : "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblsub 
      Caption         =   "SUBJECTS:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
