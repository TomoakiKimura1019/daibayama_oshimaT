Attribute VB_Name = "GpibSubRoutines"
Option Explicit

'��ܰ�ނ̒�`
Global Const GP_GTL As Long = &H1
Global Const GP_SDC As Long = &H4
Global Const GP_PPC As Long = &H5
Global Const GP_GET As Long = &H8
Global Const GP_TCT As Long = &H9
Global Const GP_LLO As Long = &H11
Global Const GP_DCL As Long = &H14
Global Const GP_PPU As Long = &H15
Global Const GP_SPE As Long = &H18
Global Const GP_SPD As Long = &H19
Global Const GP_MLA As Long = &H20
Global Const GP_UNL As Long = &H3F
Global Const GP_MTA As Long = &H40
Global Const GP_UNT As Long = &H5F

'Windows�ŋK�肳��Ă���l
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_CONTENTS = &H3

'ACX-GPIB(W32) Help�R���e�L�X�g
Global Const HLP_SAMPLES = 274
Global Const HLP_SAMPLES_BASIC = 275
Global Const HLP_SAMPLES_EVENT = 276
Global Const HLP_SAMPLES_MULTILINE = 277
Global Const HLP_SAMPLES_MULTIMETER = 278
Global Const HLP_SAMPLES_POLLING = 279
Global Const HLP_SAMPLES_PARARELL = 280
Global Const HLP_SAMPLES_VOLT = 281

'LoadProperty, SaveProperty�Ŏg�p
Global Const ERR_FILE_NOT_FOUND = 190
Global Const ERR_FILE_COULD_NOT_OPEN = 191
Global Const ERR_FILE_WRITE = 192
Global Const ERR_FILE_READ = 193
Global Const ERR_FILE_UNKNOWN = 194
Global Const ERR_FILE_INVALID_FORMAT = 195

'Win32 API�̃R�[���̂��߂̐錾
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Function GpibInit(Gp As Object, RetSts As String) As Long
'�������̂��߂̃T�u���[�`���B�I�u�W�F�N�g���������Ƃ��Ď󂯎��܂��B
'�߂�l : ����I�� = 0�A�ُ�I�� = 1

Dim ret As Long

'�ď������h�~�̂���
    Gp.Exit

    GpibInit = 0
    
'Ini�AIfc�ARen��3�̃��\�b�h�ŏ������̂ЂƂ����܂�ɂȂ�܂��B
'�{�[�h�̏�����
    ret = Gp.Ini
    If ret <> 0 Then
        GpibInit = CheckRetCode("Ini", ret, RetSts)
        Exit Function
    End If
    
'�}�X�^�̎��݈̂ȉ���2�̃��\�b�h�����s���܂�
    If Gp.MasterSlave = 0 Then
    'IFC(Interface Clear)�̑��o
        ret = Gp.Ifc
        If ret <> 0 Then
            GpibInit = CheckRetCode("Ifc", ret, RetSts)
            Exit Function
        End If
    '�����[�g���C����L���ɂ���
        ret = Gp.Ren
        If ret <> 0 Then
            GpibInit = CheckRetCode("Ren", ret, RetSts)
            Exit Function
        End If
    End If

'����I���̂Ƃ��͈ȉ��̕������Ԃ��܂�
    RetSts = "����������"
    GpibInit = 0

End Function

Public Function GpibEnd(Gp As Object, RetSts As String) As Long
'�I���̂��߂̃T�u���[�`���B�I�u�W�F�N�g���������Ƃ��Ď󂯎��܂��B

Dim ret As Long

'�}�X�^�̎��݈̂ȉ���2�̃��\�b�h�����s���܂�
    If Gp.MasterSlave = 0 Then
    '�����[�g���C���̃��Z�b�g
        ret = Gp.Resetren
        If ret <> 0 Then
            GpibEnd = CheckRetCode("Resetren", ret, RetSts)
            Exit Function
        End If
    End If

'�I�������̎��s
    ret = Gp.Exit
    If ret <> 0 Then
        GpibEnd = CheckRetCode("Exit", ret, RetSts)
        Exit Function
    End If

    RetSts = "����I��"
    GpibEnd = 0

End Function

Public Function CheckRetCode(Buf As String, RetCode As Long, RetBuf As String) As Long
'�G���[�`�F�b�N�T�u���[�`���B�\�����郁�\�b�h���Ɩ߂�l�������Ƃ��Ď󂯎��܂��B

Dim CheckRet As Long
Dim RetSts As Long
Dim TextErr As String

    TextErr = Buf + " : ����I��"
    CheckRet = RetCode And &HFF
    RetSts = 0
    
    If (CheckRet >= 3) Then
        RetSts = 1
        If (CheckRet = 3) Then TextErr = Buf + " : FIFO���ɂ܂��f�[�^���c���Ă��܂�": GoTo CheckStatus
        If (CheckRet = 80) Then TextErr = Buf + " : I/O�A�h���X�G���[": GoTo CheckStatus
        If (CheckRet = 128) Then TextErr = Buf + " : �f�[�^��M�\�萔�𒴂�����(��M)�܂���SRQ����M���Ă��܂���(�|�[�����O)": GoTo CheckStatus
        If (CheckRet = 200) Then TextErr = Buf + " : �X���b�h���쐬�ł��܂���": GoTo CheckStatus
        If (CheckRet = 240) Then TextErr = Buf + " : Esc�L�[��������܂���": GoTo CheckStatus
        If (CheckRet = 241) Then TextErr = Buf + " : File���o�̓G���[": GoTo CheckStatus
        If (CheckRet = 242) Then TextErr = Buf + " : �A�h���X�w��~�X": GoTo CheckStatus
        If (CheckRet = 243) Then TextErr = Buf + " : �o�b�t�@�w��G���[": GoTo CheckStatus
        If (CheckRet = 244) Then TextErr = Buf + " : �z��T�C�Y�G���[": GoTo CheckStatus
        If (CheckRet = 245) Then TextErr = Buf + " : �o�b�t�@�����������܂�": GoTo CheckStatus
        If (CheckRet = 246) Then TextErr = Buf + " : �s���ȃI�u�W�F�N�g���ł�": GoTo CheckStatus
        If (CheckRet = 247) Then TextErr = Buf + " : �f�o�C�X���̉��̃`�F�b�N�������ł�": GoTo CheckStatus
        If (CheckRet = 248) Then TextErr = Buf + " : �s���ȃf�[�^�^�ł�": GoTo CheckStatus
        If (CheckRet = 249) Then TextErr = Buf + " : ����ȏ�f�o�C�X��ǉ��ł��܂���": GoTo CheckStatus
        If (CheckRet = 250) Then TextErr = Buf + " : �f�o�C�X����������܂���": GoTo CheckStatus
        If (CheckRet = 251) Then TextErr = Buf + " : �f���~�^���f�o�C�X�Ԃň���Ă��܂�": GoTo CheckStatus
        If (CheckRet = 252) Then TextErr = Buf + " : GP-IB�G���[": GoTo CheckStatus
        If (CheckRet = 253) Then TextErr = Buf + " : �f���~�^�݂̂���M���܂���": GoTo CheckStatus
        If (CheckRet = 254) Then TextErr = Buf + " : �^�C���A�E�g���܂���": GoTo CheckStatus
        If (CheckRet = 255) Then TextErr = Buf + " : �p�����[�^�G���[": GoTo CheckStatus
                TextErr = Buf + " : ���̃T���v���ł̓G���[�R�[�h" & CheckRet & "�̓T�|�[�g���Ă��܂���B"
    End If

CheckStatus:
    '----- Ifc & Srq Receive Status Message ------------
    CheckRet = RetCode And &HFF00
    If (CheckRet = &H100) Then TextErr = TextErr + " -- SRQ����M���܂��� <�X�e�[�^�X>": GoTo CheckEnd
    If (CheckRet = &H200) Then TextErr = TextErr + " -- IFC����M���܂��� <�X�e�[�^�X>": GoTo CheckEnd
    If (CheckRet = &H300) Then TextErr = TextErr + " -- SRQ��IFC����M���܂��� <�X�e�[�^�X>"

CheckEnd:
    RetBuf = TextErr
    CheckRetCode = RetSts

End Function

Public Function DevidedString(Base_Str As String, Str_Cnt As Long) As String

'Base_Str�̒�����","�ŋ�؂�ꂽStr_Cnt�Ԗڂ̕������Ԃ��܂��B
'Str_Cnt=1�̎��A�擪�̕������Ԃ��܂��B�܂��AStr_Cnt�Ԗڂ̕�����
'�Ȃ������ꍇ,�܂�Str_Cnt=0,Str_Cnt>100�������ꍇ�ɂ� ""��Ԃ��܂��B
Dim StrLenPre(100) As Integer
Dim StrLenAft As Integer
Dim BaseLen As Integer
Dim Count As Integer

    If (Str_Cnt = 0) Or (Str_Cnt > 100) Or (Base_Str = "") Then
        DevidedString = ""
        Exit Function
    End If

    '�n���ꂽ������̒������擾���܂�
    BaseLen = Len(Trim$(Base_Str))
    StrLenPre(1) = 0

    For Count = 1 To Str_Cnt
        '","�̂���ʒu���擾���܂�
        StrLenAft = InStr(StrLenPre(Count) + 1, Base_Str, ",")
        If StrLenAft = 0 Then
            '�w�肳�ꂽ�ʒu������","��������Ȃ������ꍇ
            If Count = Str_Cnt Then
                '��؂�ꂽ�Ō�̈ʒu�̏ꍇ
                StrLenAft = BaseLen + 1
            Else
                '��؂�ꂽ�����A�w�肳�ꂽ�ʒu(Str_Cnt)��
                '�傫�������ꍇ(���ʂƂ���""��Ԃ��܂�)
                StrLenAft = 1
            End If
            Exit For
        End If
        '���̌����J�n�ʒu�̎w��
        StrLenPre(Count + 1) = StrLenAft
    Next

    DevidedString = Trim$(Mid$(Base_Str, StrLenPre(Str_Cnt) + 1, StrLenAft - 1))

End Function
