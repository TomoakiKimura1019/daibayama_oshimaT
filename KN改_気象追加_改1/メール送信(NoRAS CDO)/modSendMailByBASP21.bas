Attribute VB_Name = "modSendMailByBASP21"
'*******************************************************************************
'   EҰّ��M�@�\ ��BSMTP.dll(BASP21),UNLHA32.dll�K�{
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
Option Explicit
'**********���������ۂɑ��M���s�Ȃ��i�K�ł��̒l���u1�v�ɕύX���ĉ�����������
#Const cnsSW_TEST = 0       ' �e�X�g��(=0)
'#Const cnsSW_TEST = 1       ' �{��(=1)
'**********���������ۂɑ��M���s�Ȃ��i�K�ł��̒l���u1�v�ɕύX���ĉ�����������
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4
Private Const g_cnsNG = "NG"
Private Const g_cnsOK = "OK"
Private Const g_cnsYen = "\"
Private Const g_cnsLZH = ".lzh"
Private Const g_cnsERRMSG1 = "������������܂���B"
Private Const g_cnsCNT1 = 3     ' �i�[ð���1�����ڂ̗v�f��(�Œ�)
Private Const MAX_PATH = 260
' Ұّ��MAPI(BASP21)
Private Declare Function SendMail Lib "BSMTP.dll" _
    (szServer As String, szTo As String, szFrom As String, _
     szSubject As String, szBody As String, szFile As String) As String
' LHA���k�𑀍삷��API(UNLHA32)
Private Declare Function Unlha Lib "UNLHA32.dll" _
    (ByVal lhWnd As Long, ByVal szCmdLine As String, _
     ByVal szOutPut As String, ByVal wSize As Long) As Long
' �޲�ٱ��ߴ��ذ���w�肵�Đڑ�(IE4�ȏ�K�{)
Private Declare Function InternetDial Lib "WININET.dll" _
    (ByVal hwndParent As Long, ByVal lpszConnectoid As String, _
     ByVal dwFlags As Long, lpdwConnection As Long, _
     ByVal dwReserved As Long) As Long
' �޲�ٱ���ID���w�肵�Đؒf
Private Declare Function InternetHangUp Lib "WININET.dll" _
    (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
' ����ީ����ق�Ԃ�
Private Declare Function FindWindow Lib "USER32.dll" _
    Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As Any) As Long
' Sleep
Private Declare Sub Sleep Lib "KERNEL32.dll" _
    (ByVal dwMilliseconds As Long)
' SYSTEM�ިڸ�ؖ��擾API
Private Declare Function GetSystemDirectory Lib "KERNEL32.dll" _
    Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
' Windows��TEMP̫��ގ擾
Private Declare Function GetTempPath Lib "KERNEL32.dll" _
    Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'*******************************************************************************
' EҰّ��M�@�\(BSMTP.dll�K�{)
'*******************************************************************************
' [����]
'   strDialUp   : �޲�ٱ��ߓo�^��(�޲�ٱ��߂��Ȃ��������ݸ)
'   strDomain   : ��Ҳݖ�(xxxx.co.jp��)
'   strSMTP     : SMTP���ޖ�(smtp.xxxx.co.jp,mail.xxxx.co.jp��)
'   strPort     : �ʏ�͢25�,���ݸ�̏ꍇ�͢25�
'   strTimeOut  : �60��ʂ��K��,���ݸ�̏ꍇ�͢60�
'   strFromName : ���M������
'   strFromAddr : ���M�����ڽ
'   vntToName   : ���於��(�����̏ꍇ�͔z���Ă���)
'   vntToAddr   : ������ڽ(�����̏ꍇ�͔z���Ă���,�z��v�f���͈��於�̂ƈ�v������)
'   vntCCName   : CC���於��(�����̏ꍇ�͔z���Ă���)
'   vntCCAddr   : CC������ڽ(�����̏ꍇ�͔z���Ă���,�z��v�f����CC���ƈ�v������)
'   vntBCCName  : BCC���於��(�����̏ꍇ�͔z���Ă���)
'   vntBCCAddr  : BCC������ڽ(�����̏ꍇ�͔z���Ă���,�z��v�f����BCC���ƈ�v������)
'   swOwnerBCC  : True�̏ꍇ����M�����ڽ��BCC�ɉ�����
'   strSubj     : ����
'   strMessage  : �{��(�������t�����ľ��)
'   strCaption  : �e����޳��Caption
'   vntFileName : ���߽�Y�ţ�ٖ�(�����̏ꍇ�͔z���Ă���) ���Ȃ��ꍇ�����ݸ
'   strLzhFile  : ��L�Y�ţ�ق����k����ꍇ�͂��̈��ķ�ٖ�(�߽���s�v)
'   intDelMode  : ���k���̍폜���@(0=�폜�Ȃ�, 1=���ķ�ق��폜, 2=��̧�ق��폜)
' [�߂�l]
'   "OK"=����, ����ȊO�ʹװү����
'*******************************************************************************
Public Function SendMailByBASP21(strDialUp As String, _
                                 strDomain As String, _
                                 strSMTP As String, _
                                 strPort As String, _
                                 strTimeOut As String, _
                                 strFromName As String, _
                                 strFromAddr As String, _
                                 vntToName As Variant, _
                                 vntToAddr As Variant, _
                                 vntCCName As Variant, _
                                 vntCCAddr As Variant, _
                                 vntBCCName As Variant, _
                                 vntBCCAddr As Variant, _
                                 swOwnerBCC As Boolean, _
                                 strSubj As String, _
                                 strMessage As String, _
                                 Optional strCaption As String, _
                                 Optional vntFileName As Variant, _
                                 Optional strLzhFile As String, _
                                 Optional intDelMode As Integer) As String
    Dim xlAPP As Application
    Dim strDIAL_ENTRY As String     ' �޲�ٱ��ߴ��ذ
    Dim strSV_Name As String        ' ��Ҳ�/SMTP:�߰�:��ѱ��
    Dim strMailFrom As String       ' ���M���o�^
    Dim strMailto As String         ' ���M��o�^
    Dim strTable() As String        ' �z��l�i�[ð���
    Dim MAX2 As Integer             ' ð��قɊi�[�����v�f��(2�����ڍő�l)
    Dim CNT2() As Integer           ' ð��قɊi�[�����v�f��(2�����ڊe�v�f)
    Dim vntName As Variant          ' ���於Work
    Dim vntAddr As Variant          ' ���ڽWork
    Dim strPathName As String       ' �Y�ţ�ق�̫��ޖ�
    Dim strFileName As String       ' �Y�ţ��
    Dim swLine As Byte              ' �޲�ِڑ�����
    Dim hWnd As Long                ' ����޳�����
    Dim lngConnID As Long           ' �ȸ��ݺ���
    Dim IX As Long                  ' ð���Index
    Dim IX1 As Long                 ' ð���Index
    Dim IX2 As Long                 ' ð���Index
    Dim IX3 As Long                 ' ð���Index
    Dim lngRet As Long              ' ���ݺ���
    Dim strRC As String             ' BASP21�߂�l
    Dim strMSG As String            ' ү����
    Dim vntMSG As Variant           ' ү����Work
    Dim strName As String           ' Work
    Dim strAddr As String           ' Work

    SendMailByBASP21 = g_cnsNG
    Set xlAPP = Application
    
    ' BSMTP.dll�̑��݊m�F
    If Dir(FP_GET_SYSTEM_PATH & "BSMTP.dll", vbNormal) = "" Then
        SendMailByBASP21 = _
            "���M�R���|�[�l���g�BSMTP.dll����C���X�g�[������Ă��܂���B"
        Exit Function
    End If
    
'-------------------------------------------------------------------------------
' ����������(�����n���p�����[�^�̍쐬)
    
    ' ��Ҳ�/SMTP:�߰�:��ѱ��
    If Trim$(strPort) = "" Then strPort = "25"
    If Trim$(strTimeOut) = "" Then strTimeOut = "60"
    strSV_Name = Trim$(strDomain) & "/" & _
                 Trim$(strSMTP) & ":" & _
                 Trim$(strPort) & ":" & _
                 Trim$(strTimeOut)
    
    ' Variant���ږ��g�p�̏ꍇ�̑Ή�
    If IsError(vntToName) Then vntToName = ""
    If IsError(vntToAddr) Then vntToAddr = ""
    If IsError(vntCCName) Then vntCCName = ""
    If IsError(vntCCAddr) Then vntCCAddr = ""
    If IsError(vntBCCName) Then vntBCCName = ""
    If IsError(vntBCCAddr) Then vntBCCAddr = ""
    If IsError(vntFileName) Then vntFileName = ""
                 
    ' ���M���o�^
    If Trim$(strFromAddr) = "" Then
        SendMailByBASP21 = "���M���̃��[���A�h���X������܂���"
        Exit Function
    End If
    If Trim$(strFromName) = "" Then
        strMailFrom = Trim$(strFromAddr)
    Else
        strMailFrom = Trim$(strFromName) & _
            "<" & Trim$(strFromAddr) & ">"
    End If
    
    ' �z��ň����n�����\�������鍀�ڂ�S�ĕ�ð��قɊi�[������
    ' (�Ȍ�͑S�Ĕz��̕�����ϐ��Ƃ��ď����ł���)
    MAX2 = 0
    ReDim strTable(g_cnsCNT1, MAX2)
    ReDim CNT2(g_cnsCNT1)
    vntMSG = Array("����", "CC����", "BCC����", "�Y�t�t�@�C����")
    For IX1 = 0 To g_cnsCNT1
        Select Case IX1
            Case 0: vntName = vntToName:   vntAddr = vntToAddr      ' ����
            Case 1: vntName = vntCCName:   vntAddr = vntCCAddr      ' CC
            Case 2: vntName = vntBCCName:  vntAddr = vntBCCAddr     ' BCC
            Case 3: vntName = vntFileName: vntAddr = vntFileName    ' �Y�ţ��
        End Select
        IX3 = 0
        If IsArray(vntAddr) = True Then
            ' �i�[ð��قɔz����i�[
            For IX2 = LBound(vntAddr) To UBound(vntAddr)
                On Error GoTo MakeArray_ARRAY2
                strAddr = Trim$(vntAddr(IX2))
                If ((IX1 < g_cnsCNT1) And (IX2 <= UBound(vntName))) Then
                    On Error GoTo MakeArray_ARRAY3
                    strName = Trim$(vntName(IX2))
                Else
                    strName = ""
                End If
                GoSub MakeArray_SUB
            Next IX2
        Else
            If IX1 < g_cnsCNT1 Then
                strName = Trim$(vntName)
            Else
                strName = ""
            End If
            strAddr = Trim$(vntAddr)
            GoSub MakeArray_SUB
        End If
        CNT2(IX1) = IX3
    Next IX1
    If CNT2(0) < 1 Then
        SendMailByBASP21 = "����̃��[���A�h���X������܂���"
        Exit Function
    End If
    
    ' ���M�҂�BCC�ɒǉ��w��̏���(swOwnerBCC�w��̏ꍇ)
    If swOwnerBCC = True Then
        CNT2(2) = CNT2(2) + 1
        If CNT2(2) > MAX2 Then
            MAX2 = CNT2(2)
            ReDim Preserve strTable(g_cnsCNT1, MAX2)
        End If
        If strFromName <> "" Then
            strTable(2, CNT2(2)) = strFromName & "<" & strFromAddr & ">"
        Else
            strTable(2, CNT2(2)) = strFromAddr
        End If
    End If
    
    ' ���M��o�^(����,CC,BCC��Tab��؂�÷�Ăɂ���)
    strMailto = ""
    For IX1 = 0 To 2
        If CNT2(IX1) >= 1 Then
            Select Case IX1
                Case 1: strMailto = strMailto & vbTab & "cc"
                Case 2: strMailto = strMailto & vbTab & "bcc"
            End Select
            IX = 1
            Do While IX <= CNT2(IX1)
                ' 2���ڈȍ~��Tab��؂�ž��
                If strMailto = "" Then
                    strMailto = strTable(IX1, IX)
                Else
                    strMailto = strMailto & vbTab & strTable(IX1, IX)
                End If
                IX = IX + 1
            Loop
        End If
    Next IX1
    
    ' �Y�ţ�ُ���
    strFileName = ""
    If Trim$(strLzhFile) <> "" Then
        ' ���ķ�ق��w�肳��Ă���ꍇ�͈��ķ�ق�Y�ţ�قɎw��(�P��̧��)
        strMSG = FP_ArchiveByUNLHA32(strLzhFile, vntFileName, strCaption)
        If strMSG <> g_cnsOK Then
            SendMailByBASP21 = "���k�t�@�C���̍쐬�Ɏ��s���܂����B" & vbCr & _
                strMSG
            Exit Function
        End If
        strFileName = strLzhFile
    Else
        ' ���ķ�ق��w�肳��Ă��Ȃ��ꍇ��Tab��؂�÷�Ăɂ���
        IX1 = g_cnsCNT1
        IX = 1
        Do While IX <= CNT2(IX1)
            If strFileName = "" Then
                strFileName = strTable(IX1, IX)
            Else
                strFileName = strFileName & vbTab & strTable(IX1, IX)
            End If
            IX = IX + 1
        Loop
    End If
    
'-------------------------------------------------------------------------------
' �����M����
    
    ' �޲�ٱ��ߐڑ�(���ؖ����w�肳��Ă���ꍇ�̂�)
    strDIAL_ENTRY = Trim$(strDialUp)
    If strDIAL_ENTRY <> "" Then
        strMSG = ""
        swLine = 0
        xlAPP.StatusBar = "�" & strDIAL_ENTRY & "��ɐڑ����ł��D�D�D�D"
        ' ����޳����ق��擾
        hWnd = FP_GET_HWND(strCaption)
        lngConnID = 0
        ' �ӰĐڑ����N��
        lngRet = InternetDial(hWnd, strDIAL_ENTRY, _
            INTERNET_AUTODIAL_FORCE_UNATTENDED, lngConnID, 0&)
        If ((lngRet <> 0) And (lngRet <> 633)) Then
            strMSG = "�" & strDIAL_ENTRY & "��ւ̐ڑ��Ɏ��s���܂����B"
            Select Case lngRet
                Case 623: strMSG = strMSG & vbCr & "�@(�޲�ٴ��ذ�����s����)"
                Case 668: strMSG = strMSG & vbCr & "�@(�߽ܰ�ނ����o�^)"
                Case Else: strMSG = strMSG & vbCr & _
                    "�@(���̑��װ : " & CStr(lngRet) & " )"
            End Select
            SendMailByBASP21 = strMSG
            Exit Function
        End If
        swLine = 1
    End If
    
    ' BASP21(BSMTP.dll)���s
    xlAPP.StatusBar = "���[���𑗐M���ł��D�D�D�D"
    On Error GoTo BASP_ERROR
#If cnsSW_TEST = 1 Then
    ' �{��
    strRC = SendMail(strSV_Name, strMailto, strMailFrom, strSubj, _
        strMessage, strFileName)
#Else
    ' �e�X�g(�����\���̂�)
    MsgBox "�E��Ҳ�/SMTP:�߰�:��ѱ�� = " & strSV_Name & vbCr & _
           "�E���� = " & strMailto & vbCr & _
           "�E���o�l = " & strMailFrom & vbCr & _
           "�E���� = " & strSubj & vbCr & _
           "�E�Y�t = " & strFileName & vbCr & vbCr & _
           "������̓e�X�g�p�̊m�F���b�Z�[�W�ł��B" & vbCr & _
           "�@�{�Ԃɐ؂�ւ���ɂ́AmodSendMailByBASP21_2�̍ŏ��ɂ���" & vbCr & _
           "�@�R���p�C���X�C�b�`�ucnsSW_TEST�v�̒l���u1�v�ɕύX���ĕۑ����ĉ������B"
#End If
    
    ' �޲�ٺȸ��݂�ؒf
    If swLine = 1 Then
        ' �����ؒf
        xlAPP.StatusBar = "�" & strDIAL_ENTRY & "���ؒf���ł��D�D�D�D"
        InternetHangUp lngConnID, 0&
        AppActivate xlAPP.Caption       ' Excel��è�ނɂ���
        swLine = 0
    End If
    
    If strRC <> "" Then
        SendMailByBASP21 = strRC & vbCr & vbCr & _
            "�T�[�o�[�ɐڑ��ł��Ȃ����A�ؒf����܂����B"
        xlAPP.StatusBar = False
        Exit Function
    End If
    
'-------------------------------------------------------------------------------
' ���I������(���ķ�َw�莞�̎���폜����)
    
    ' ���ķ�ق��쐬�����ꍇ�͍폜���邩���肷��(���M���펞�̂�)
    If strLzhFile <> "" Then
        xlAPP.DisplayAlerts = False
        Select Case intDelMode
            Case 1
                ' ���ķ�ق��폜����
                Kill strLzhFile
            Case 2
                ' ��̧�ق��폜����
                If IsArray(vntFileName) = True Then
                    ' �z��w�莞�͏����폜
                    vntAddr = vntFileName
                    For IX2 = LBound(vntAddr) To UBound(vntAddr)
                        On Error GoTo MakeArray_ARRAY2
                        strAddr = Trim$(vntAddr(IX2))
                        On Error Resume Next
                        Kill strAddr
                    Next IX2
                    On Error GoTo 0
                Else
                    ' �P��̧�َw��
                    strFileName = Trim$(vntFileName)
                    On Error Resume Next
                    Kill strFileName
                    On Error GoTo 0
                End If
        End Select
        xlAPP.DisplayAlerts = True
    End If
    
    SendMailByBASP21 = g_cnsOK
    AppActivate xlAPP.Caption           ' Excel��è�ނɂ���
    xlAPP.StatusBar = False
    Exit Function

'-------------------------------------------------------------------------------
' 1�����Q�ƂŴװ�̏ꍇ��2�����Ƃ��ď���(�͈ٔ͊i�[�Ή�)
MakeArray_ARRAY2:
    On Error GoTo MakeArray_ERROR
    strAddr = Trim$(vntAddr(IX2, 1))
    Resume Next

'-------------------------------------------------------------------------------
' 1�����Q�ƂŴװ�̏ꍇ��2�����Ƃ��ď���(�͈ٔ͊i�[�Ή�)
MakeArray_ARRAY3:
    On Error GoTo MakeArray_ERROR
    strName = Trim$(vntName(IX2, 1))
    Resume Next

'-------------------------------------------------------------------------------
' �i�[ð��قɾ�Ă���
MakeArray_SUB:
    If strAddr <> "" Then
        IX3 = IX3 + 1
        If IX3 > MAX2 Then
            ' �ő�l�Ŋi�[ð��ق̗v�f����ύX
            MAX2 = IX3
            ReDim Preserve strTable(g_cnsCNT1, MAX2)
        End If
        If strName <> "" Then
            strTable(IX1, IX3) = strName & "<" & strAddr & ">"
        Else
            strTable(IX1, IX3) = strAddr
        End If
        ' ����ɑ��M�ұ��ڽ������ꍇ��BCC�t�����Ȃ�
        If strAddr = strFromAddr Then swOwnerBCC = False
    End If
    Return

'-------------------------------------------------------------------------------
' �װ����
MakeArray_ERROR:
    SendMailByBASP21 = "�p�����[�^�o�^�����Ɏ��s���܂����B(" & _
        vntMSG(IX1) & ")" & vbCr & "  (" & Err.Description & ")"
    xlAPP.StatusBar = False
    Exit Function

'-------------------------------------------------------------------------------
' BASP21���s���װ
BASP_ERROR:
    strRC = "���[�����M�R���|�[�l���g�BASP21������s�ł��܂���B" & _
        vbCr & Err.Number & " " & Err.Description
    Resume Next
    
End Function

'*******************************************************************************
' ̧�و��k�@�\(UNLHA32.dll�K�{)
'*******************************************************************************
' [����]
'   strTarget   : ���k���̧�ٖ�
'   vntSource   : ���k�Ώۂ�̧�ٖ�(�����̏ꍇ�͔z���Ă���)
'   strCaption  : �e����޳��Caption
' [�߂�l]
'   "OK"=����, ����ȊO�ʹװү����
'*******************************************************************************
Public Function FP_ArchiveByUNLHA32(strTarget As String, _
                                    vntSource As Variant, _
                                    Optional strCaption As String) As String
    Dim xlAPP As Application
    Dim strFileName As String       ' ̧�ٖ�(work)
    Dim strPathName As String       ' ���ķ�ق�̫���
    Dim strExeName As String        ' �����𓀈��ķ��
    Dim strCommand As String        ' UNLHA�����ײ�
    Dim strBuffer As String         ' Work
    Dim strEXT As String            ' �g���q
    Dim IX As Long                  ' ð���Index
    Dim hWnd As Long                ' ����޳�����
    Dim strMSG As String            ' UNLHA32�װү����
    Dim cntSource As Long           ' ����̧�ِ�
    
    FP_ArchiveByUNLHA32 = g_cnsNG
    Set xlAPP = Application
    xlAPP.StatusBar = "���k�t�@�C���쐬���D�D�D�D"
    
    ' UNLHA32.dll�̑��݊m�F
    If Dir(FP_GET_SYSTEM_PATH & "UNLHA32.dll", vbNormal) = "" Then
        FP_ArchiveByUNLHA32 = _
            "���k�R���|�[�l���g�UNLHA32����C���X�g�[������Ă��܂���B"
        Exit Function
    End If
    
    ' �o��̧�ق��߽���Ȃ��ꍇ��TEMP̫��ނɏo��
    strFileName = Trim$(strTarget)
    If ((Left$(strFileName, 2) <> "\\") And _
        (Mid$(strFileName, 2, 2) <> ":\")) Then
        ' TEMP̫��ނ����
        strPathName = FP_GET_TEMP_PATH
        strTarget = strPathName & strFileName
    Else
        strTarget = strFileName
        IX = Len(strFileName)
        Do While IX > 1
            If Mid$(strFileName, IX, 1) = g_cnsYen Then Exit Do
            IX = IX - 1
        Loop
        strPathName = Left$(strFileName, IX)
    End If
    
    ' �g���q�̔���
    strEXT = StrConv(Right$(strTarget, 3), vbUpperCase)
    If ((strEXT <> "LZH") And (strEXT <> "EXE")) Then
        If Mid$(strTarget, Len(strTarget) - 3, 1) <> "." Then
            strTarget = strTarget & g_cnsLZH
        Else
            strFileName = Left$(strTarget, Len(strTarget) - 4)
            strTarget = strFileName & g_cnsLZH
        End If
    ElseIf strEXT = "EXE" Then
        strExeName = strTarget
        strFileName = Left$(strTarget, Len(strTarget) - 4)
        strTarget = strFileName & g_cnsLZH
    End If
    
    ' ����̧�ق����݂���ꍇ�͍폜
    If Dir(strTarget, vbNormal) <> "" Then Kill strTarget
    
    ' UNLHA�̺����ײ݂�ҏW
    strCommand = "a """ & strTarget & """"
    If IsArray(vntSource) = True Then
        For IX = LBound(vntSource) To UBound(vntSource)
            On Error GoTo UnLha_ARRAY
            strFileName = Trim$(vntSource(IX))
            On Error GoTo 0
            If strFileName <> "" Then
                If GetAttr(strFileName) And vbDirectory Then
                    FP_ArchiveByUNLHA32 = _
                        "�����w��ł̓t�H���_�͎w��ł��܂���B"
                    GoTo UnLha_EXIT
                Else
                    strCommand = strCommand & " """ & strFileName & """"
                End If
                cntSource = cntSource + 1
            End If
        Next IX
    ElseIf IsError(vntSource) <> True Then
        strFileName = Trim$(vntSource)
        If strFileName <> "" Then
            If GetAttr(strFileName) And vbDirectory Then
                ' ̫��ގw��̏ꍇ�͔z���S�Ă��i�[
                strCommand = strCommand & " -d1 """ & _
                    Left(strFileName, Len(strFileName) - _
                        Len(Dir(strFileName, vbDirectory))) & "\"" " & _
                    Dir(strFileName, vbDirectory)
            Else
                strCommand = strCommand & " """ & strFileName & """"
            End If
            cntSource = cntSource + 1
        End If
    End If
    
    ' �L���ȓ���̧�ق��Ȃ��ꍇ�͖���
    If cntSource < 1 Then
        strTarget = ""
        FP_ArchiveByUNLHA32 = g_cnsOK
        Exit Function
    End If
    
    On Error GoTo UnLha_ERROR
    ' ����޳����ق��擾
    hWnd = FP_GET_HWND(strCaption)
    ' �����ײ݂ɏ]����UNLHA�𑀍�
    strBuffer = String(256, Chr$(0))
    If Unlha(hWnd, strCommand, strBuffer, Len(strBuffer)) = 0& Then
        If strEXT = "EXE" Then
            ' EXE�`���w��̏ꍇ�͎����𓀏��ɂɕϊ�
            strCommand = "s -gw2 """ & strTarget & """ """ & strPathName & """"
            strBuffer = String(256, Chr$(0))
            If Unlha(hWnd, strCommand, strBuffer, Len(strBuffer)) = 0& Then
                Kill strTarget
                strTarget = strExeName
                FP_ArchiveByUNLHA32 = g_cnsOK
            Else
                FP_ArchiveByUNLHA32 = _
                    Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
            End If
        Else
            FP_ArchiveByUNLHA32 = g_cnsOK
        End If
    Else
        FP_ArchiveByUNLHA32 = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
    End If
    GoTo UnLha_EXIT

'-------------------------------------------------------------------------------
' �z�񑀍�װ�Ή�(2�����z��̏ꍇ�͍Ĕz�u���Ė߂�)
UnLha_ARRAY:
    On Error GoTo UnLha_ERROR2
    strFileName = Trim$(vntSource(IX, 1))
    Resume Next
    
'-------------------------------------------------------------------------------
' UNLHA32���s���װ
UnLha_ERROR:
    FP_ArchiveByUNLHA32 = "���k�R���|�[�l���g�UNLHA32������s�ł��܂���B" & _
        vbCr & Err.Number & " " & Err.Description
    GoTo UnLha_EXIT

'-------------------------------------------------------------------------------
' �z�񑀍쎞�װ
UnLha_ERROR2:
    FP_ArchiveByUNLHA32 = "���̓t�@�C���w�肪����������܂���B(UNLHA32)" & _
        vbCr & Err.Number & " " & Err.Description

'-------------------------------------------------------------------------------
' UNLHA32�����I��
UnLha_EXIT:
    On Error Resume Next
    AppActivate xlAPP.Caption
End Function

'*******************************************************************************
' Windows��SYSTEM�t�H���_�擾
'*******************************************************************************
' [�߂�l] SYSTEM̫���(�װ����)
'*******************************************************************************
Private Function FP_GET_SYSTEM_PATH() As String
    Dim strBuffer As String
    Dim strPathName As String
    
    ' Buffer���m��
    strBuffer = String(MAX_PATH, Chr(0))
    ' SYSTEM�ިڸ�ؖ��擾
    Call GetSystemDirectory(strBuffer, MAX_PATH)
    ' Null�����̎�O�܂ł�L���Ƃ��ĕ\��(��������ݸ�̧�ٖ��ϊ���)
    strPathName = Left$(strBuffer, InStr(1, strBuffer, Chr(0)) - 1)
    If Right$(strPathName, 1) <> g_cnsYen Then strPathName = strPathName & g_cnsYen
    FP_GET_SYSTEM_PATH = strPathName
End Function

'*******************************************************************************
' Windows��TEMP�t�H���_�擾
'*******************************************************************************
' [�߂�l] TEMP̫���(�װ����)
'*******************************************************************************
Private Function FP_GET_TEMP_PATH() As String
    Dim strBuffer As String
    Dim strPathName As String
    
    ' Buffer���m��
    strBuffer = String(MAX_PATH, Chr(0))
    ' SYSTEM�ިڸ�ؖ��擾
    Call GetTempPath(MAX_PATH, strBuffer)
    ' Null�����̎�O�܂ł�L���Ƃ��ĕ\��(��������ݸ�̧�ٖ��ϊ���)
    strPathName = Left$(strBuffer, InStr(1, strBuffer, Chr(0)) - 1)
    If Right$(strPathName, 1) <> g_cnsYen Then strPathName = strPathName & g_cnsYen
    FP_GET_TEMP_PATH = strPathName
End Function

'*******************************************************************************
' �E�B���h�E�n���h���̎擾
'*******************************************************************************
' [����]
'   strCaption  : ����޳��Caption(�׽�͎������f)
' [�߂�l]
'   hWnd        : ����޳����ْl(���s�;��)
'*******************************************************************************
Private Function FP_GET_HWND(strCaption As String) As Long
    Dim strClassName As String
    
    strClassName = "XLMAIN"
    Select Case strCaption
        Case "": strCaption = Application.Caption
        Case Application.Caption
        Case Else
            ' UserForm�̏ꍇ
            If Val(Application.Version) <= 8 Then
                strClassName = "ThunderXFrame"      ' Excel97
            Else
                strClassName = "ThunderDFrame"      ' Excel2000�ȍ~
            End If
    End Select
    On Error GoTo GET_HWND_ERR
    FP_GET_HWND = FindWindow(strClassName, strCaption)
    Exit Function

'-------------------------------------------------------------------------------
GET_HWND_ERR:
    FP_GET_HWND = 0&
End Function

'-----------------------------<< End of Source >>-------------------------------
