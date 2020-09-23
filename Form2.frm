VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirm Overwrite"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1920
      Width           =   1305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   285
      Left            =   2453
      TabIndex        =   2
      Top             =   1230
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   285
      Left            =   833
      TabIndex        =   1
      Top             =   1230
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   173
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   105
      Width           =   4185
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Kill Form1.GetFilePath(Form1.Text11.Text) & "Integrator.ico"
Form1.Command1Secondary
Unload Me
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
Form1.btnFlat Command1
Form1.btnFlat Command2
End Sub
Private Sub Text1_GotFocus()
Text2.SetFocus
End Sub
