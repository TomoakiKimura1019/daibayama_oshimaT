VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "åvë™íÜâÊñ "
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   6510
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   1  'êÖïΩ
      TabIndex        =   3
      Text            =   "Form2.frx":0CCA
      Top             =   2760
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'óºï˚
      TabIndex        =   0
      Text            =   "Form2.frx":0CD0
      Top             =   1000
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      Caption         =   "*** åv ë™ íÜ ***"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   18
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    'Call KijyunIn(3)
End Sub

Private Sub Command2_Click()
'    MDY = Now
'    Tensu% = 3
'    HeikinKaisuu = 3
'    Call KIJYUN_READ
'    Call SOKUTEI("master.dat")
End Sub

Private Sub Form_Load()
    Text1.Text = ""
'shira    Call GTS8init
End Sub



