Attribute VB_Name = "modCDOmail"
'*******************************************************************************
'   CDO�Ń��[���𑗐M����   ���Q�Ɛݒ��
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
'   [�Q�Ɛݒ�]
'   �EMicrosoft CDO for Windows 2000 Library
'     (or Microsoft CDO for Exchange 2000 Library)
'*******************************************************************************
Option Explicit

'**********���������ۂɑ��M���s�Ȃ��i�K�ł��̒l���u1�v�ɕύX���ĉ�����������
'#Const cnsSW_TEST = 0       ' �e�X�g��(=0)
#Const cnsSW_TEST = 1       ' �{��(=1)
'**********���������ۂɑ��M���s�Ȃ��i�K�ł��̒l���u1�v�ɕύX���ĉ�����������
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4
Private Const g_cnsNG = "NG"
Private Const g_cnsOK = "OK"
Private Const g_cnsYen = "\"
Private Const g_cnsERRMSG1 = "������������܂���B"
Private Const g_cnsCNT1 = 3     ' �i�[ð���1�����ڂ̗v�f��(�Œ�)
Private Const MAX_PATH = 260

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

'*******************************************************************************
' ���[�����M(CDO)  ���Q�Ɛݒ��
'*******************************************************************************
' [����]
'  �@MailSmtpServer : SMTP�T�[�o��(����IP�A�h���X)
'  �AMailFrom       : ���M���A�h���X
'  �BMailTo         : ����A�h���X(�����̏ꍇ�̓J���}�ŋ�؂�)
'  �CMailCc         : CC�A�h���X(�����̏ꍇ�̓J���}�ŋ�؂�)
'  �DMailBcc        : BCC�A�h���X(�����̏ꍇ�̓J���}�ŋ�؂�)
'  �EMailSubject    : ����
'  �FMailBody       : �{��(���s��vbCrLf�t��)
'  �GMailAddFile    : �Y�t�t�@�C��(�����̏ꍇ�̓J���}�ŋ�؂邩�z��n��) ��Option
'  �HMailCharacter  : �����R�[�h�w��(�f�t�H���g��Shift-JIS)              ��Option
' [�߂�l]
'  ���펞�F"OK", �G���[���F"NG"+�G���[���b�Z�[�W
'*******************************************************************************
Public Function SendMailCDO(strDialUp As String, MailSmtpServer As String, MailFrom As String, MailTo As String, MailCc As String, MailBcc As String, MailSubject As String, MailBody As String, Optional MailAddFile As Variant, Optional MailCharacter As String)
    Const cnsOK = "OK"
    Const cnsNG = "NG"
    Dim objCDO As New CDO.Message
    Dim vntFILE As Variant
    Dim IX As Long
    Dim strCharacter As String, strBody As String, strChar As String
    
'    On Error GoTo SendMailByCDO_ERR
    SendMailCDO = cnsNG
    
    Dim strDIAL_ENTRY As String     ' �޲�ٱ��ߴ��ذ
    Dim strMSG As String            ' ү����
    Dim swLine As Byte              ' �޲�ِڑ�����
    Dim hWnd As Long                ' ����޳�����
    Dim lngConnID As Long           ' �ȸ��ݺ���
    Dim lngRet As Long              ' ���ݺ���
    
'    ' �޲�ٱ��ߐڑ�(���ؖ����w�肳��Ă���ꍇ�̂�)
'    strDIAL_ENTRY = Trim$(strDialUp)
'    If strDIAL_ENTRY <> "" Then
'        strMSG = ""
'        swLine = 0
'        MainForm.StatusBar1.Panels(1) = "�" & strDIAL_ENTRY & "��ɐڑ����ł��D�D�D�D"
'        ' ����޳����ق��擾
'        hWnd = MainForm.hWnd 'FP_GET_HWND(strCaption)
'        lngConnID = 0
'        ' �ӰĐڑ����N��
'        lngRet = InternetDial(hWnd, strDIAL_ENTRY, INTERNET_AUTODIAL_FORCE_UNATTENDED, lngConnID, 0&)
'        If ((lngRet <> 0) And (lngRet <> 633)) Then
'            strMSG = "�" & strDIAL_ENTRY & "��ւ̐ڑ��Ɏ��s���܂����B"
'            Select Case lngRet
'                Case 623: strMSG = strMSG & vbCr & "�@(�޲�ٴ��ذ�����s����)"
'                Case 668: strMSG = strMSG & vbCr & "�@(�߽ܰ�ނ����o�^)"
'                Case Else: strMSG = strMSG & vbCr & "�@(���̑��װ : " & CStr(lngRet) & " )"
'            End Select
'            SendMailCDO = strMSG
'            Exit Function
'        End If
'        swLine = 1
'    Else
'        Exit Function
'    End If
    
    ' �����R�[�h�w��̊m�F
    If MailCharacter <> "" Then
        ' �w�肠��̏ꍇ�͎w��l���Z�b�g
        strCharacter = MailCharacter
    Else
        ' �w��Ȃ��̏ꍇ��Shift-JIS�Ƃ���
        strCharacter = cdoShift_JIS
    End If
    
    ' �{���̉��s�R�[�h�̊m�F
    ' Lf�݂̂̏ꍇCr+Lf�ɕϊ�
    strBody = Replace(MailBody, vbLf, vbCrLf)
    ' ��L�Ō���Cr+Lf�̏ꍇCr+Cr+Lf�ɂȂ�̂�Cr+Lf�ɖ߂�
    MailBody = Replace(strBody, vbCr & vbCrLf, vbCrLf)
    
    With objCDO
'        With .Configuration.Fields                          ' �ݒ荀��
'            .Item(cdoSendUsingMethod) = cdoSendUsingPort    ' �O��SMTP�w��
'            .Item(cdoSMTPServer) = MailSmtpServer           ' SMTP�T�[�o��
'            .Item(cdoSMTPServerPort) = 25                   ' �|�[�g��
'            .Item(cdoSMTPConnectionTimeout) = 60            ' �^�C���A�E�g
'            .Item(cdoSMTPAuthenticate) = cdoAnonymous       ' 0
'            .Item(cdoLanguageCode) = strCharacter           ' �����Z�b�g�w��
'            .Update                                         ' �ݒ���X�V
'        End With
        'SMTP�F�؂Ȃ炱����
'        strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
        With .Configuration.Fields
            .Item(cdoSendUsingMethod) = 2               ' �O��SMTP�w��
            .Item(cdoSMTPServer) = MailSmtpServer      ' SMTP�T�[�o��
            .Item(cdoSMTPServerPort) = 465            ' �|�[�g��
            .Item(cdoSMTPUseSSL) = True               ' SSL���g���ꍇ��True
            .Item(cdoSMTPAuthenticate) = 1            ' 1(Basic�F��)/2�iNTLM�F�؁j
            .Item(cdoSendUserName) = "atic.alertmail@gmail.com"
            .Item(cdoSendPassword) = "idappe99"
            .Item(cdoSMTPConnectionTimeout) = 60          ' �^�C���A�E�g
            .Item(cdoLanguageCode) = strCharacter           ' �����Z�b�g�w��
            .Update
        End With
        .Fields("urn:schemas:mailheader:X-Mailer") = "CDO mail"
        .Fields("urn:schemas:mailheader:Importance") = "High"
        .Fields("urn:schemas:mailheader:Priority") = 1
        .Fields("urn:schemas:mailheader:X-Priority") = 1
        .Fields("urn:schemas:mailheader:X-MsMail-Priority") = "High"
        .Fields.Update
        
        .MimeFormatted = True
        .Fields.Update
        .From = MailFrom                        ' ���M��
        .To = MailTo                            ' ����
        If MailCc <> "" Then .CC = MailCc       ' CC
        If MailBcc <> "" Then .BCC = MailBcc    ' BCC
        .Subject = MailSubject                  ' ����
        .TextBody = MailBody                    ' �{��
        .TextBodyPart.Charset = strCharacter    ' �����Z�b�g�w��(�{��)
        ' �Y�t�t�@�C���̓o�^(�����Ή�)
        If ((VarType(MailAddFile) <> vbError) And (VarType(MailAddFile) <> vbBoolean) And (VarType(MailAddFile) <> vbEmpty) And (VarType(MailAddFile) <> vbNull)) Then
            If IsArray(MailAddFile) Then
                For IX = LBound(MailAddFile) To UBound(MailAddFile)
                    .AddAttachment MailAddFile(IX)
                Next IX
            ElseIf MailAddFile <> "" Then
                vntFILE = Split(CStr(MailAddFile), ",")
                For IX = LBound(vntFILE) To UBound(vntFILE)
                    If Trim(vntFILE(IX)) <> "" Then
                        .AddAttachment Trim(vntFILE(IX))
                    End If
                Next IX
            End If
        End If
        .Send                                   ' ���M
    End With
    Set objCDO = Nothing
    SendMailCDO = cnsOK

    ' �޲�ٺȸ��݂�ؒf
    If swLine = 1 Then
        ' �����ؒf
        MainForm.StatusBar1.Panels(1) = "�" & strDIAL_ENTRY & "���ؒf���ł��D�D�D�D"
        InternetHangUp lngConnID, 0&
        swLine = 0
    End If
Dim strRC
    If strRC <> "" Then
        SendMailCDO = strRC & vbCr & vbCr & "�T�[�o�[�ɐڑ��ł��Ȃ����A�ؒf����܂����B"
        MainForm.StatusBar1.Panels(1) = False
        Exit Function
    End If




Exit Function

'-------------------------------------------------------------------------------
SendMailByCDO_ERR:
    SendMailCDO = cnsNG & Err.Number & " " & Err.Description
    On Error Resume Next
    Set objCDO = Nothing
End Function
