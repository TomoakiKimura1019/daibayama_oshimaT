VERSION 5.00
Begin VB.Form frmInitMsg 
   Caption         =   "メッセージ"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   600
   ScaleWidth      =   4125
   StartUpPosition =   2  '画面の中央
   Begin VB.Label Label1 
      Caption         =   "*** 変位計の初期化中 ***"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInitMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



