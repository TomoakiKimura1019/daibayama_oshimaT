VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmTELset 
   Caption         =   "�ً}���A�������"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   Icon            =   "frmTELset.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Frame Frame2 
      Caption         =   "�k ���� �l"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   12975
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "�l�r ����"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   12735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ۑ�"
      Height          =   495
      Index           =   0
      Left            =   13440
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   495
      Index           =   1
      Left            =   13440
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Z���ҏW"
      Height          =   2175
      Left            =   13320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "�R�s�["
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�؂���"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "�\��t��"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
   End
   Begin VCF150Ctl.F1Book F1Book1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5530
      _0              =   $"frmTELset.frx":08CA
      _1              =   $"frmTELset.frx":0CD3
      _2              =   $"frmTELset.frx":10DC
      _3              =   $"frmTELset.frx":14E5
      _4              =   $"frmTELset.frx":18EE
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "frmTELset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean
Dim maxFLD As Integer
Dim L(500) As String, co As Integer
Dim DM(500) As String, DMco As Integer
Dim Tabl_path As String
Public CurrentDIR As String   'keisoku.exe������t�H���_

Sub InitSheet()
    Dim f As Integer, L As String
    Dim i As Integer, j As Integer
    Dim Dco As Long, Dsw As Boolean
    Dim maxREC As Long, Rco As Long, Rpst1 As Long, Rpst2 As Long, Rsw As Boolean
    Dim FLDno As Integer
    Dim FILENAME As String

    On Error Resume Next
    Screen.MousePointer = 11
    
    F1Book1.DoSafeEvents = True
    F1Book1.EnableProtection = True
    
    F1Book1.MaxCol = 7
    F1Book1.ColText(1) = "��Ж��P"
    F1Book1.ColText(2) = "��Ж��Q"
    F1Book1.ColText(3) = "��Ж��R"
    F1Book1.ColText(4) = "�d�b�ԍ��P"
    F1Book1.ColText(5) = "�d�b�ԍ��Q"
    F1Book1.ColText(6) = "�S���Ґ�"
    F1Book1.ColText(7) = "�S���ҕ�"
    
    FILENAME = Tabl_path & "�ً}�A��.dat"
    If Dir(FILENAME) <> "" Then
        f = FreeFile
        Open FILENAME For Input Shared As #f
            i = 0: DMco = 0
            Do While Not (EOF(f))
                Line Input #f, L
                If Left$(L, 1) = ":" Then
                    DMco = DMco + 1
                    DM(DMco) = L
                Else
                    i = i + 1
                    F1Book1.TextRC(i, 1) = Trim(SEEKmoji(L, 5, 30))
                    F1Book1.TextRC(i, 2) = Trim(SEEKmoji(L, 35, 30))
                    F1Book1.TextRC(i, 3) = Trim(SEEKmoji(L, 65, 30))
                    F1Book1.TextRC(i, 4) = Trim(SEEKmoji(L, 95, 20))
                    F1Book1.TextRC(i, 5) = Trim(SEEKmoji(L, 115, 20))
                    F1Book1.TextRC(i, 6) = Trim(SEEKmoji(L, 135, 16))
                    F1Book1.TextRC(i, 7) = Trim(SEEKmoji(L, 151, 16))
                    
                End If
            Loop
        Close #f
    End If

    F1Book1.SetSelection 1, 1, F1Book1.MaxRow, F1Book1.MaxCol
    F1Book1.SetProtection False, True
    F1Book1.EnableProtection = True
    F1Book1.SetSelection i + 1, 1, i + 1, 1
    F1Book1.ShowActiveCell
    
    F1Book1.Modified = False
    F1Book1.DoSafeEvents = True
    F1Book1.SetFocus
    
    Screen.MousePointer = 0
End Sub

Sub SaveData()
    Dim i As Integer, f As Integer
    Dim j As Integer
    Dim SS As String
    Dim Incel As Integer
    
    f = FreeFile
    Open Tabl_path & "�ً}�A��.dat" For Output Lock Write As #f
        For i = 1 To DMco
            Print #f, DM(i)
        Next i
        
        For i = 1 To F1Book1.MaxRow
            SS = F1Book1.TextRC(i, 1) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 1), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 2) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 2), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 3) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 3), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 4) & Space$(20 - LenB(StrConv(F1Book1.TextRC(i, 4), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 5) & Space$(20 - LenB(StrConv(F1Book1.TextRC(i, 5), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 6) & Space$(16 - LenB(StrConv(F1Book1.TextRC(i, 6), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 7) & Space$(16 - LenB(StrConv(F1Book1.TextRC(i, 7), vbFromUnicode)))
            If Trim(SS) <> "" Then Print #f, Format(i, "!@@@@") & SS
        Next i
    Close #f
    
    F1Book1.Modified = False
End Sub

Private Sub Command1_Click(Index As Integer)
'    On Error Resume Next
    
    If Index = 0 Then
        Call SaveData
        MsgBox "�ۑ����������܂����B", vbInformation
        F1Book1.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub Command3_Click(Index As Integer)
On Error Resume Next
    
    If Index = 0 Then F1Book1.EditCopy: F1Book1.EditClear F1ClearValues '�؂���
    If Index = 1 Then F1Book1.EditCopy        '�R�s�[
    If Index = 2 Then F1Book1.EditPasteValues '�\��t��
    F1Book1.SetFocus
End Sub

Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
'    On Error Resume Next
'    Select Case F1Book1.Col
'    Case 1
'        If IsDate(EditString) = True Then
'            F1Book1.TextRC(F1Book1.Row, 1) = Format(DateValue(EditString), "yyyy/mm/dd")
'            F1Book1.CancelEdit
'            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
'        Else
'            MsgBox "���t�ƔF���ł��Ȃ��l�����͂���܂����B������x�A���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
'            F1Book1.CancelEdit
'        End If
'    Case 2
'        If IsDate(EditString) = True Then
'            F1Book1.TextRC(F1Book1.Row, 2) = Format(TimeValue(EditString), "h:nn")
'            F1Book1.CancelEdit
'            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
'        Else
'            MsgBox "���ԂƔF���ł��Ȃ��l�����͂���܂����B������x�A���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
'            F1Book1.CancelEdit
'        End If
'    Case Else
'        If IsNumeric(EditString) = False Then
'            MsgBox "���l�ƔF���ł��Ȃ��l�����͂���܂����B������x�A���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
'            F1Book1.CancelEdit
'        End If
'        If F1Book1.TextRC(F1Book1.Row, 1) = "" Or F1Book1.TextRC(F1Book1.Row, 2) = "" Then
'            MsgBox "�v�����t�E�v�����Ԃ����͂���Ă��܂���B��ɁA�v�����t�E�v�����Ԃ���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
'            F1Book1.CancelEdit
'        End If
'    End Select
End Sub

Private Sub F1Book1_SelChange()
    Dim SS As String
        
    Select Case F1Book1.Col
    Case 1 To 3
        SS = "��Ж�����͂��܂���ő啶�����ͤ�S�p������15�����܂łł��"
    Case 4 To 5
        SS = "�d�b�ԍ�����͂��܂���ő啶�����ͤ�S�p������10�����܂łł��"
    Case Else
        SS = "�S���҂���͂��܂���ő啶�����ͤ�S�p������8�����܂łł��"
    End Select
    Label1.Caption = SS
End Sub

Private Sub Form_Load()
    ChDrive App.Path
    ChDir App.Path
    
    CurrentDIR = App.Path
    If Right(CurrentDIR, 1) = "\" Then Else CurrentDIR = CurrentDIR & "\"
    
    Tabl_path = GetIni("�t�H���_��", "���f�[�^", CurrentDIR & "�A���ݒ�.ini")
    Call InitSheet
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer
    If F1Book1.Modified Then
        Response = MsgBox("�ύX���ۑ�����Ă��܂���B�ۑ����܂����H", vbYesNoCancel + vbExclamation, "�I���̊m�F")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then Call SaveData
    End If
End Sub
'**********************************************************************************************
'   �t�H�[���̃T�C�Y��ύX�����ꍇ�ɃR���g���[���̈ʒu��ݒ�
'**********************************************************************************************
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Frame1
        .Left = ScaleWidth - 1560
        .Top = 120
    End With
    With Frame2
        .Left = 120
        .Top = ScaleHeight - 1125
    End With
    
    With Command1(0)
        .Left = ScaleWidth - 1440
        .Top = 2760
    End With
    With Command1(1)
        .Left = ScaleWidth - 1440
        .Top = ScaleHeight - 645
    End With
    
    With F1Book1
        .Left = 120
        .Top = 120
        .Height = Me.ScaleHeight - 1350
        .Width = Me.ScaleWidth - 1905
    End With
'Debug.Print Me.Width, Me.Height
End Sub


'**********************************************************************************************
'   �����񂩂�w�肵�����������̕������Ԃ��܂��B
'   �i�S�p�������Q�����A���p�������P�����Ƃ��܂��B�j
'**********************************************************************************************
Public Function SEEKmoji(strCheckString As String, mojiST As Integer, mojiMAX As Integer) As String

    'For�J�E���^
    Dim i As Long
    
    '�����Ώە�����̒������i�[
    Dim lngCheckSize As Long
    
    'ANSI�ւ̕ϊ���̕������i�[
    Dim lngANSIStr As Long
    
    Dim co As Integer '������
    Dim SS As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: SS = ""
    For i = 1 To lngCheckSize
        'StrConv��Unicode����ANSI�ւƕϊ�
        lngANSIStr = LenB(StrConv(Mid(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            SS = SS + Mid(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = SS
End Function

