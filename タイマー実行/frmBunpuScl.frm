VERSION 5.00
Begin VB.Form frmBunpuScl 
   Caption         =   "���z�}�X�P�[���ݒ�"
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
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.CommandButton Command1 
      Caption         =   "��ݾ�"
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
      Caption         =   "�P�ڐ���l"
      BeginProperty Font 
         Name            =   "�l�r �o����"
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
         Alignment       =   1  '�E����
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
        '�G���[�`�F�b�N
        If IsNumeric(Text1(0).Text) = False Then
            MsgBox "���l�ƔF���ł��Ȃ��l�����͂���܂����B������x�A���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
            Exit Sub
        End If
        If Text1(0).Text = 0 Then
            MsgBox "0���傫�����l����͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
            Exit Sub
        End If
        
        Bunpu.Xscl = Text1(0).Text
        Call WriteIni("�f�ʕ��z�}�ݒ�", "�P�ڐ���l", CStr(Bunpu.Xscl), CurrentDIR & "�v���ݒ�.ini")
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
