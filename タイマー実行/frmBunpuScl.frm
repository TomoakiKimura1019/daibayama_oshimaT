VERSION 5.00
Begin VB.Form frmBunpuScl 
   Caption         =   "分布図スケール設定"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   Icon            =   "frmBunpuScl.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "１目盛り値"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      Begin VB.TextBox Text1 
         Alignment       =   1  '右揃え
         Height          =   270
         Index           =   0
         Left            =   600
         MaxLength       =   5
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmBunpuScl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public OWARI As Boolean

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        'エラーチェック
        If IsNumeric(Text1(0).Text) = False Then
            MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            Exit Sub
        End If
        If Text1(0).Text = 0 Then
            MsgBox "0より大きい数値を入力してください。", vbCritical, "エラーメッセージ"
            Exit Sub
        End If
        
        Bunpu.Xscl = Text1(0).Text
        Call WriteIni("断面分布図設定", "１目盛り値", CStr(Bunpu.Xscl), CurrentDIR & "計測設定.ini")
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.top = frmBunpu.top + 500
    Me.Left = frmBunpu.Left + 500
    
    Text1(0).Text = Bunpu.Xscl
    
    frmCLOSE.bunpuScl = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmCLOSE.bunpuScl = True
End Sub
