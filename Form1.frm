VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "T.A.S. Internet Explorer Toolbar Integrator"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   5205
   End
   Begin VB.PictureBox Picture5 
      Height          =   360
      Left            =   1500
      Picture         =   "Form1.frx":27A2
      ScaleHeight     =   300
      ScaleWidth      =   525
      TabIndex        =   49
      Top             =   5250
      Width           =   585
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   6825
      TabIndex        =   2
      Top             =   1290
      Width           =   510
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   6855
      TabIndex        =   1
      Top             =   720
      Width           =   525
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5715
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   720
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5610
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "C:\"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5130
      Left            =   0
      TabIndex        =   3
      Top             =   15
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   9049
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Create New"
      TabPicture(0)   =   "Form1.frx":2AAC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label21"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Check1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command6"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Picture1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Picture4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text11"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text13"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text12"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Edit Existing"
      TabPicture(1)   =   "Form1.frx":2AC8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command12"
      Tab(1).Control(1)=   "Command11"
      Tab(1).Control(2)=   "Command10"
      Tab(1).Control(3)=   "Picture3"
      Tab(1).Control(4)=   "Picture2"
      Tab(1).Control(5)=   "Combo1"
      Tab(1).Control(6)=   "Text10"
      Tab(1).Control(7)=   "Text9"
      Tab(1).Control(8)=   "Text8"
      Tab(1).Control(9)=   "Text7"
      Tab(1).Control(10)=   "Text6"
      Tab(1).Control(11)=   "Text5"
      Tab(1).Control(12)=   "Text1"
      Tab(1).Control(13)=   "Command7"
      Tab(1).Control(14)=   "Command8"
      Tab(1).Control(15)=   "Command9"
      Tab(1).Control(16)=   "Label19"
      Tab(1).Control(17)=   "Label18"
      Tab(1).Control(18)=   "Label17"
      Tab(1).Control(19)=   "Label16"
      Tab(1).Control(20)=   "Label15"
      Tab(1).Control(21)=   "Label14"
      Tab(1).Control(22)=   "Label13"
      Tab(1).Control(23)=   "Label12"
      Tab(1).Control(24)=   "Label11"
      Tab(1).ControlCount=   25
      Begin VB.CommandButton Command12 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71593
         TabIndex        =   54
         Top             =   4710
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4313
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1950
         Width           =   630
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   428
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1950
         Width           =   3840
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   428
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2445
         Width           =   3855
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   428
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1245
         Width           =   3810
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4515
         Picture         =   "Form1.frx":2AE4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   46
         Top             =   615
         Width           =   480
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72969
         TabIndex        =   45
         Top             =   4710
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Save Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74371
         TabIndex        =   44
         Top             =   4710
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -70425
         Picture         =   "Form1.frx":2DEE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   41
         Top             =   585
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -74730
         Picture         =   "Form1.frx":30F8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         Top             =   585
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   375
         Picture         =   "Form1.frx":3402
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         Top             =   600
         Width           =   480
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3803
         TabIndex        =   23
         Top             =   3615
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Use Icon In Exicutable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1755
         TabIndex        =   22
         Top             =   1590
         Value           =   1  'Checked
         Width           =   2160
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2768
         TabIndex        =   21
         Top             =   3615
         Width           =   960
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   443
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2970
         Width           =   4485
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1095
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   750
         Width           =   3180
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4328
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2445
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4283
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1245
         Width           =   660
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Set Visible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2085
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3345
         Value           =   1  'Checked
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate And Install"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   503
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3615
         Width           =   2190
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74040
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   690
         Width           =   3450
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74872
         TabIndex        =   13
         Text            =   "Text10"
         Top             =   1230
         Width           =   4965
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74857
         TabIndex        =   12
         Text            =   "Text9"
         Top             =   4215
         Width           =   4965
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74857
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   3705
         Width           =   4965
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74857
         TabIndex        =   10
         Text            =   "Text7"
         Top             =   3225
         Width           =   4305
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74857
         TabIndex        =   9
         Text            =   "Text6"
         Top             =   2745
         Width           =   4305
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74857
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   2250
         Width           =   4305
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74872
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1740
         Width           =   4965
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70507
         TabIndex        =   6
         Top             =   2265
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70507
         TabIndex        =   5
         Top             =   2745
         Width           =   600
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Open"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70507
         TabIndex        =   4
         Top             =   3225
         Width           =   600
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hot Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4320
         TabIndex        =   48
         Top             =   435
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   195
         TabIndex        =   47
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hot Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -70575
         TabIndex        =   43
         Top             =   375
         Width           =   825
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74970
         TabIndex        =   42
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tools Menu Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -73447
         TabIndex        =   39
         Top             =   4020
         Width           =   2265
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Bar Tool Tip Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -73365
         TabIndex        =   38
         Top             =   3525
         Width           =   2115
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hot Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -73155
         TabIndex        =   37
         Top             =   3045
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73185
         TabIndex        =   36
         Top             =   2550
         Width           =   1740
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Exicutable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73170
         TabIndex        =   35
         Top             =   2055
         Width           =   1725
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UUID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -73170
         TabIndex        =   34
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Button Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73200
         TabIndex        =   33
         Top             =   1050
         Width           =   1770
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Bar Tool Tip Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1590
         TabIndex        =   32
         Top             =   2790
         Width           =   2205
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hot Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2310
         TabIndex        =   31
         Top             =   2250
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Icon"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2235
         TabIndex        =   30
         Top             =   1770
         Width           =   900
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Main Exicutable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1980
         TabIndex        =   29
         Top             =   1065
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Button Title"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1958
         TabIndex        =   28
         Top             =   555
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   435
         TabIndex        =   27
         Top             =   4140
         Width           =   4500
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Generated Tool Bar UUID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1568
         TabIndex        =   26
         Top             =   3960
         Width           =   2235
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tool Bar Tool Tip Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   -73260
         TabIndex        =   25
         Top             =   2730
         Width           =   2205
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CoCreateGuid Requirements
  'Windows NT/2000/XP: Requires Windows NT 3.1 or later.
  'Windows 95/98: Requires Windows 95 or later.
  'Header: Declared in objbase.h.
  'Library: Use ole32.lib.

'StringFromGUID2 Requirements
  'Windows NT/2000/XP: Requires Windows NT 3.1 or later.
  'Windows 95/98: Requires Windows 95 or later.
  'Header: Declared in objbase.h.
  'Library: Use ole32.lib.

'Windows 95/98/Me:SetWindowLongW is supported by the Microsoft Layer for Unicode (MSLU).
'SetWindowLongA is also supported to provide more consistent behavior across all Windows operating systems.
'To use these versions, you must add certain files to your application, as outlined in Microsoft Layer for Unicode on Windows 95/98/Me Systems.
'http://www.microsoft.com/downloads/details.aspx?FamilyId=73BA7BD7-ED06-4F0D-80A4-2A7EEAEE17E2&displaylang=en

Option Explicit
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private Type GUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(8) As Byte
End Type
Private LastSelectedMain As String
Private LastSelectedHot As String
Private FadeIn As Integer
Private Sub Check2_Click()
Text4.SetFocus
If Check2.Value = 1 Then
Command3.Enabled = False
Command4.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Text12.Text = Text11.Text
Text13.Text = Text11.Text
PicLoad
Else
Command3.Enabled = True
Command4.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Text13.Text = LastSelectedHot
Text12.Text = LastSelectedMain
PicLoad
End If
End Sub
Private Sub Combo1_Click()
Text4.SetFocus
Text1.Text = List2.List(Combo1.ListIndex)
FetchSettings
End Sub
Private Sub Command1_Click()
Text4.SetFocus
If Check2.Value = 1 Then
 If FileExists(GetFilePath(Text11.Text) & "Integrator.ico") Then
 Form2.Show
 Exit Sub
 Else
 End If
End If
Command1Secondary
End Sub
Public Sub Command1Secondary()
If Check2.Value = 1 Then
 SavePicture Picture1.Picture, GetFilePath(Text11.Text) & "Integrator.ico"
End If
If Text2.Text = "" Then
MsgBox "You Must Fill In The Tool Button Title Box!"
Exit Sub
ElseIf Text3.Text = "" Then
MsgBox "You Must Fill In The Tool Bar Tool Tip Text Box!"
Exit Sub
ElseIf Text11.Text = "" Then
MsgBox "You Must Choose A Main Program File !"
Exit Sub
ElseIf Text12.Text = "" And Check2.Value = 0 Then
MsgBox "You Must Choose A Main Icon File !"
Exit Sub
ElseIf Text13.Text = "" And Check2.Value = 0 Then
MsgBox "You Must Choose A Hot Icon File !"
Exit Sub
End If
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
CreateToolBarTool
End Sub
Private Function GenerateUUID() As String
Dim udtGUID As GUID
Dim strGUID As String
Dim bytGUID() As Byte
Dim lngLen As Long
Dim lngRetVal As Long
Dim lngPos As Long
lngLen = 40
bytGUID = String(lngLen, 0)
CoCreateGuid udtGUID
lngRetVal = StringFromGUID2(udtGUID, VarPtr(bytGUID(0)), lngLen)
strGUID = bytGUID
If (Asc(Mid$(strGUID, lngRetVal, 1)) = 0) Then
    lngRetVal = lngRetVal - 1
End If
strGUID = Left$(strGUID, lngRetVal)
GenerateUUID = strGUID
End Function
Private Sub CreateToolBarTool()
Dim TheUUID As String
TheUUID = GenerateUUID
Label10.Caption = TheUUID
CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "ButtonText", Text2.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "CLSID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
If Check1.Value = 1 Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "Default Visible", "Yes"
Else
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "Default Visible", "No"
End If
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "Exec", Text11.Text
If Check2.Value = 1 Then
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "HotIcon", GetFilePath(Text11.Text) & "Integrator.ico"
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "Icon", GetFilePath(Text11.Text) & "Integrator.ico"
Else
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "HotIcon", Text13.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "Icon", Text12.Text
End If
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "MenuStatusBar", Text3.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & TheUUID, "MenuText", Text2.Text
End Sub

Private Sub Command10_Click()
Text4.SetFocus
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "ButtonText", Text10.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "Exec", Text5.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "Icon", Text6.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "HotIcon", Text7.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "MenuStatusBar", Text8.Text
SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "MenuText", Text9.Text
PopulateEditor
End Sub

Private Sub Command11_Click()
Dim DUUID As String
DUUID = Text1.Text
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "ButtonText"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "CLSID"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Default Visible"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "HotIcon"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "MenuStatusBar"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "MenuText"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Exec"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Icon"
'DeleteKey dont seem to work in XP however removing the values above does remove the button.
'I guess users in NT will just have to use a reg cleaner.
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text
PopulateEditor
End Sub
Private Sub Command12_Click()
Unload Me
End Sub
Private Sub Command2_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.FileName = ""
CommonDialog1.DialogTitle = "Choose main tool file to be opened !"
CommonDialog1.Filter = "Exicutable files|*.EXE;*.BAT|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text11.Text = CommonDialog1.FileName
Text2.Text = Left(GetFileName(CommonDialog1.FileName), Len(GetFileName(CommonDialog1.FileName)) - 4)
Text3.Text = Text2.Text
If Check2.Value = 1 Then
Text12.Text = Text11.Text
Text13.Text = Text11.Text
PicLoad
End If
OpenProblem:
End Sub
Private Sub Command3_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.FileName = ""
CommonDialog1.DialogTitle = "Choose main icon to be opened !"
CommonDialog1.Filter = "Icon files|*.ICO|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text12.Text = CommonDialog1.FileName
LastSelectedMain = Text12.Text
PicLoad
OpenProblem:
End Sub
Private Sub Command4_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.DialogTitle = "Choose hot icon file to be opened !"
CommonDialog1.Filter = "Icon files|*.ICO|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text13.Text = CommonDialog1.FileName
LastSelectedHot = Text13.Text
PicLoad
OpenProblem:
End Sub
Private Sub Command5_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Label10.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
End Sub
Private Sub Command6_Click()
Dim DUUID As String
DUUID = Label10.Caption
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "ButtonText"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "CLSID"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Default Visible"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "HotIcon"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "MenuStatusBar"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "MenuText"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Exec"
DeleteValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & DUUID, "Icon"
'DeleteKey dont seem to work in XP however removing the values above does remove the button.
'I guess users in NT will just have to use a reg cleaner.
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Label10.Caption
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = False
Command6.Enabled = False
Label10.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
End Sub
Private Sub Command7_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.FileName = Text5.Text
CommonDialog1.DialogTitle = "Choose main tool file to be opened !"
CommonDialog1.Filter = "Exicutable files|*.EXE;*.BAT|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text5.Text = CommonDialog1.FileName
OpenProblem:
End Sub
Private Sub Command8_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.FileName = Text6.Text
CommonDialog1.DialogTitle = "Choose main icon to be opened !"
CommonDialog1.Filter = "Icon files|*.ICO|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text6.Text = CommonDialog1.FileName
GetIconFromFileIndex Text6.Text, "0", Form1.Picture2
OpenProblem:
End Sub
Private Sub Command9_Click()
Text4.SetFocus
On Error GoTo OpenProblem
CommonDialog1.FileName = Text7.Text
CommonDialog1.DialogTitle = "Choose hot icon file to be opened !"
CommonDialog1.Filter = "Icon files|*.ICO|"
CommonDialog1.FilterIndex = 1
CommonDialog1.Action = 1
Text7.Text = CommonDialog1.FileName
GetIconFromFileIndex Text7.Text, "0", Form1.Picture3
OpenProblem:
End Sub
Private Sub Form_Load()
If App.PrevInstance Then End
'************************
'Used so that in Win 95/98 form effects are not enabled.
'************************
Select Case GetVersion()
Case WIN98 'Windows 95/98
 Me.Visible = True
Case WINNT 'Windows NT
 'Me.Visible = True
 Timer1.Enabled = True
End Select
'************************
btnFlat Command1
btnFlat Command2
btnFlat Command3
btnFlat Command4
btnFlat Command5
btnFlat Command6
btnFlat Command7
btnFlat Command8
btnFlat Command9
btnFlat Command10
btnFlat Command11
btnFlat Command12
Command5.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Command6.Enabled = False
SSTab1.Tab = 0
CommonDialog1.InitDir = "C:\"
End Sub
Function btnFlat(Button As CommandButton)
SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
Public Function GetFilePath(FileNamePath As String) As String
    On Error GoTo FunctionError:
    Dim X
    Dim tString As String
    For X = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, X, 1)
    If tString = "\" Then
            GetFilePath = Left(FileNamePath, X)
            Exit Function
        End If
    Next X
FunctionError:
    GetFilePath = -1
End Function
Public Function GetFileName(file As String) As String
        Dim m
    Dim GetChr0 As String
    Dim GetChr1 As String
    For m = 1 To Len(file)
        GetChr0 = Right(file, m)
        GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then
        GetFileName = Right(GetChr0, m - 1): Exit Function
    End If
    Next m
End Function
Private Function RemoveIconFromFilename(Input1 As String) As String
    Dim m As Integer
    Dim GetChr0 As String
    Dim GetChr1 As String
    For m = 1 To Len(Input1)
        GetChr0 = Left(Input1, m)
        GetChr1 = Right(GetChr0, 1)
        If GetChr1 = "," Then
        RemoveIconFromFilename = Left(GetChr0, Len(GetChr0) - 1): Exit Function
    End If
    Next m
End Function
Private Sub FetchSettings()
Text5.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "Exec")
Text6.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "Icon")
Text7.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "HotIcon")
Text8.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "MenuStatusBar")
Text9.Text = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & Text1.Text, "MenuText")
Text10.Text = Combo1
If GetFileExtension(Text6.Text) = "ico" Then
Form1.Picture2 = LoadPicture(Text6.Text)
GetIconFromFileIndex Text6.Text, "0", Form1.Picture2
Else
GetIconFromFileIndex RemoveIconFromFilename(Text6.Text), "0", Form1.Picture2
End If
If GetFileExtension(Text7.Text) = "ico" Then
Form1.Picture3 = LoadPicture(Text7.Text)
GetIconFromFileIndex Text7.Text, "0", Form1.Picture3
Else
GetIconFromFileIndex RemoveIconFromFilename(Text7.Text), "0", Form1.Picture3
End If
End Sub
Private Sub PopulateEditor()
Dim KeyCollection As Collection
Dim KeyCollection2 As Collection
Dim Object As Variant
List1.Clear
List2.Clear
Combo1.Clear
Set KeyCollection = EnumRegistryKeys("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions")
For Each Object In KeyCollection
List1.AddItem Object
Next
Dim DButtonTitle As String
Dim X As Integer
For X = 0 To List1.ListCount - 1
DButtonTitle = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\" & List1.List(X), "ButtonText")
If DButtonTitle <> "Error" And DButtonTitle <> "Offline" Then
Combo1.AddItem DButtonTitle
List2.AddItem List1.List(X)
End If
Next X
Combo1.Text = Combo1.List(0)
Text1.Text = List2.List(0)
FetchSettings
End Sub
Private Sub SSTab1_GotFocus()
Text4.SetFocus
End Sub
Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SSTab1.Tab = 0 Then PopulateEditor
End Sub
Public Function GetFileExtension(FileName As String)
    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(FileName, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Function
    Else
        TempStr = Right(FileName, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Function
        Else
            TempStr = Right(FileName, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Function
            Else
                TempStr = Right(FileName, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
    
End Function
Private Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #1
        Close #1
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function
Private Sub PicLoad()
If Text12.Text > "" And GetFileExtension(Text12.Text) = "ico" Then
Form1.Picture1 = LoadPicture(Text12.Text)
ElseIf Text12.Text > "" Then
GetIconFromFileIndex Text12.Text, "0", Form1.Picture1
Else
Form1.Picture1 = Picture5.Picture
End If
If Text13.Text > "" And GetFileExtension(Text13.Text) = "ico" Then
Form1.Picture4 = LoadPicture(Text13.Text)
ElseIf Text13.Text > "" Then
GetIconFromFileIndex Text13.Text, "0", Form1.Picture4
Else
Form1.Picture4 = Picture5.Picture
End If
End Sub
Private Sub Text11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text4.SetFocus
End Sub
Private Sub Text12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text4.SetFocus
End Sub
Private Sub Text13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text4.SetFocus
End Sub
Private Sub Timer1_Timer()
If FadeIn = 255 Then Timer1.Enabled = False ': MsgBox "Timer Off !"  'Used to debug timer
MakeTransparent Me.hwnd, FadeIn
FadeIn = FadeIn + 3
If Me.Visible = False Then Me.Visible = True
End Sub
