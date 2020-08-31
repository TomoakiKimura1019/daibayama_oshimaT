Attribute VB_Name = "SyokiModule"
Option Explicit

Public MDY
Public Const MAX_CH  As Integer = 27   '      �V     �̃f�[�^��
Public CurrentDir As String   'keisoku.exe������t�H���_
Public Tabl_path As String

Public SyokiSW As Boolean
Public PoSW As Boolean

'���ݒ�
Public Const Maxkou As Integer = 1     '�v�����ڐ�
Public Const Maxdan As Integer = 1     '�v���f�ʐ�
Public Type table1
    dan   As Integer     '�f�ʔԍ�
    kou   As Integer     '���ڔԍ�
    ten   As Integer     '���_�ԍ�
    FLD   As Integer     '�f�[�^�t�@�C���t�B�[���h�ʒu
    ch    As Integer     '�v���`�����l��
    Syo   As Double      '�����l
    Kei   As Double      '�W��
    deep  As Single      '�[�x
    HAN   As String * 8  '�}��
    keiji As Integer     '�o���ω��}��}���_ 1.���� 0.���Ȃ�
End Type
Public Tbl(Maxkou, Maxdan, 10) As table1

Public Type RsInit1
    DeviceNo As Integer
    SpdNO As Integer
    PrtNO As Integer
    sizeNO As Integer
    stopNo As Integer
    Stime As Long
    Rtime As Long
End Type
Public RsInit As RsInit1


Public Type Init1
    Serch As String
    Wait As String
    PoFILE0 As String
    PoFILE1 As String
    Tensu As Integer
    HeikinKaisuu As Integer
    x0 As Double
    y0 As Double
    z0 As Double
    MH As Double
    x1 As Double
    y1 As Double
    z1 As Double
    AZIMUTH As Single
'    CO As Integer
'    AvgCO As Integer
'    Kx As Double
'    Ky As Double
'    Kz As Double
'    Kmh As Double
'    Bx As Double
'    By As Double
'    Bz As Double
'    HOKO As Single
End Type
Public InitDT As Init1

Public Type Po1
    H(2, 16) As Double
    V(2, 16) As Double
    s(2, 16) As Double
'    Hdt As Double
'    Vdt As Double
'    Sdt As Double
End Type
Public PoDT As Po1
Public RsctlFrm As Object
'�t�H�[����������ɁA�f�[�^�̍Đݒ�E�ĕ`������邽�߂̃L�[���[�h
Public Type frm1
    setTABLE As Boolean
    setKanri As Boolean
    keijiSet As Boolean
    bunpuScl As Boolean
    sinHosei As Boolean
    StartFrm As Boolean
    MSGfrm As Boolean
End Type
Public frmCLOSE As frm1

Public DH(17) As Double
Public DV(17) As Double
Public XD(20) As Double '������W
Public YD(20) As Double
Public ZD(20) As Double
Public xo(16) As Double '������W�i�O�l�j
Public yo(16) As Double
Public zo(16) As Double

Public Const RAD As Double = 3.14159265358979 / 180#



Sub Main()
    Dim rc As Integer
    Dim i As Integer
    
    If App.PrevInstance = True Then
        MsgBox "���ɋN�����Ă��܂��B", vbCritical, "�N���G���["
        End
    End If
    
'    ChDrive App.Path
'    ChDir App.Path
    SetCurrentDirectory App.Path
    
    CurrentDir = App.Path
    If Right(CurrentDir, 1) = "\" Then Else CurrentDir = CurrentDir & "\"
    
    Set RsctlFrm = frmPOset
    
    Tabl_path = GetIni("�t�H���_��", "���f�[�^", CurrentDir & "�v���ݒ�.ini")
    Call ReadTabl
    
    '�v������
    InitDT.Serch = GetIni("�v������", "�T�[�`�ݒ�", CurrentDir & "�v���ݒ�.ini")
    InitDT.Wait = GetIni("�v������", "�E�F�C�g���Ԑݒ�", CurrentDir & "�v���ݒ�.ini")
    InitDT.PoFILE0 = GetIni("�v������", "���_�ʒu�t�@�C�������l", CurrentDir & "�v���ݒ�.ini")
    InitDT.PoFILE1 = GetIni("�v������", "���_�ʒu�t�@�C���O��l", CurrentDir & "�v���ݒ�.ini")
    
    InitDT.Tensu = CInt(GetIni("�v������", "���_��", CurrentDir & "�v���ݒ�.ini"))
    InitDT.x0 = CDbl(GetIni("�v������", "��B�_X", CurrentDir & "�v���ݒ�.ini"))
    InitDT.y0 = CDbl(GetIni("�v������", "��B�_Y", CurrentDir & "�v���ݒ�.ini"))
    InitDT.z0 = CDbl(GetIni("�v������", "��B�_Z", CurrentDir & "�v���ݒ�.ini"))
    InitDT.MH = CDbl(GetIni("�v������", "��B�_MH", CurrentDir & "�v���ݒ�.ini"))
    InitDT.x1 = CDbl(GetIni("�v������", "�㎋�_X", CurrentDir & "�v���ݒ�.ini"))
    InitDT.y1 = CDbl(GetIni("�v������", "�㎋�_Y", CurrentDir & "�v���ݒ�.ini"))
    InitDT.z1 = CDbl(GetIni("�v������", "�㎋�_Z", CurrentDir & "�v���ݒ�.ini"))
    InitDT.AZIMUTH = CDbl(GetIni("�v������", "�����p", CurrentDir & "�v���ݒ�.ini"))
    
'    i = FileCheck(InitDT.PoFILE, "�v���ʒu�t�@�C��")
'    If i = 0 Then End
   
    '�ʐM�ݒ�
    RsInit.DeviceNo = CInt(GetIni("�ʐM�ݒ�", "�ʐM�|�[�g", CurrentDir & "�v���ݒ�.ini"))
    RsInit.SpdNO = CInt(GetIni("�ʐM�ݒ�", "�ʐM���x", CurrentDir & "�v���ݒ�.ini"))
    RsInit.PrtNO = CInt(GetIni("�ʐM�ݒ�", "�p���e�B", CurrentDir & "�v���ݒ�.ini"))
    RsInit.sizeNO = CInt(GetIni("�ʐM�ݒ�", "�f�[�^��", CurrentDir & "�v���ݒ�.ini"))
    RsInit.stopNo = CInt(GetIni("�ʐM�ݒ�", "�X�g�b�v", CurrentDir & "�v���ݒ�.ini"))
    RsInit.Rtime = CLng(GetIni("�ʐM�ݒ�", "��M�^�C���A�E�g", CurrentDir & "�v���ݒ�.ini"))
    RsInit.Stime = CLng(GetIni("�ʐM�ݒ�", "���M�^�C���A�E�g", CurrentDir & "�v���ݒ�.ini"))
    
    frmSyokiset.Show

End Sub

Public Function FileCheck(FileName As String, FileTitle As String) As Integer
    Dim i As Integer

    On Error Resume Next

    i = 0
    If Dir$(FileName) = "" Then Else i = 1
    If i = 0 Then
        MsgBox FileTitle & "�t�@�C��(" & FileName & ")��������܂���B�m�F���Ă��������B", vbCritical, "�G���[���b�Z�[�W"
    End If
    
    FileCheck = i
    
    On Error GoTo 0

End Function

Public Function SEEKmoji(strCheckString As String, mojiST As Integer, mojiMAX As Integer) As String

    'For�J�E���^
    Dim i As Long
    '�����Ώە�����̒������i�[
    Dim lngCheckSize As Long
    'ANSI�ւ̕ϊ���̕������i�[
    Dim lngANSIStr As Long
    
    Dim co As Integer '������
    Dim ss As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: ss = ""
    For i = 1 To lngCheckSize
        'StrConv��Unicode����ANSI�ւƕϊ�
        lngANSIStr = LenB(StrConv(Mid$(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            ss = ss + Mid$(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = ss
End Function


'**********************************************************************************************
'   �����l�E�W���E����فE��}���_�Ȃǂ�TABLE.set
'**********************************************************************************************
Public Sub ReadTabl()
    Dim f As Integer
    Dim i As Integer
    Dim L As String
    Dim d_ID As Integer
    Dim k_ID As Integer, t_ID As Integer
'''    Dim FLDno As Integer
    
    Erase Tbl
    
    i = FileCheck(Tabl_path & "CTABLE.DAT", "���f�[�^")
    If i = 0 Then End
    
    d_ID = 1
    i = 0
    f = FreeFile
    Open Tabl_path & "CTABLE.DAT" For Input Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L
        If Left$(L, 1) <> ";" Then
            i = i + 1
            k_ID = CInt(Mid$(L, 5, 4))
            t_ID = CInt(Mid$(L, 9, 4))
            
            Tbl(k_ID, d_ID, t_ID).kou = k_ID
            Tbl(k_ID, d_ID, t_ID).ten = t_ID
            Tbl(k_ID, d_ID, t_ID).FLD = CInt(Mid$(L, 1, 4))
            Tbl(k_ID, d_ID, t_ID).ch = CInt(Mid$(L, 13, 4))
            Tbl(k_ID, d_ID, t_ID).Syo = CSng(Mid$(L, 17, 10))
            Tbl(k_ID, d_ID, t_ID).Kei = CDbl(Mid$(L, 27, 10))
            Tbl(k_ID, d_ID, t_ID).HAN = Trim$(SEEKmoji(L, 37, 8))
            
            Tbl(k_ID, d_ID, 0).ten = Tbl(k_ID, d_ID, 0).ten + 1
            
        End If
    Loop
    Close #f
End Sub


