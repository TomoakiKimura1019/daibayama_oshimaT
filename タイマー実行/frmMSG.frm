VERSION 5.00
Begin VB.Form frmMSG 
   Caption         =   "���_�G���[���b�Z�[�W"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   StartUpPosition =   1  '��Ű ̫�т̒���
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "�{�� ���� �v���Y�����m�F���Ă��������B"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "�@�B �������� ��_���ُ�ł��B"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

