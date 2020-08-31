VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'なし
   Caption         =   "データ処理"
   ClientHeight    =   14940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14940
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Skip"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   14160
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   8790
      Left            =   960
      Picture         =   "frmStart.frx":0000
      Top             =   2040
      Width           =   17070
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      BorderStyle     =   6  '実線 (ふちどり)
      Height          =   14700
      Left            =   120
      Top             =   120
      Width           =   18940
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      BackColor       =   &H00FFFFFF&
      Caption         =   "札幌ドーム屋根変位計測プログラム"
      BeginProperty Font 
         Name            =   "ＭＳ 明朝"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   18975
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      Height          =   14920
      Left            =   0
      Top             =   0
      Width           =   19180
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 15360 - 420  '11000 '16590
    Me.Width = 19200 '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
'    Me.Height = 12000 '16590
'    Me.Width = 16000 '21210
'    Left = (Screen.Width - Me.Width) / 2
'    Top = 0
    frmCLOSE.StartFrm = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmCLOSE.StartFrm = True
End Sub


