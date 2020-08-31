Attribute VB_Name = "modBSMTP"
Option Explicit

'
' �Q�Ɛݒ��BSMTP�Ƀ`�F�b�N������
'
'------------------------------------------------------
Private Declare Function SendMail Lib "bsmtp" _
      (szServer As String, szTo As String, szFrom As String, _
      szSubject As String, szBody As String, szFile As String) As String
Private Declare Function RcvMail Lib "bsmtp" _
      (szServer As String, szUser As String, szPass As String, _
      szCommand As String, szDir As String) As Variant
Private Declare Function ReadMail Lib "bsmtp" _
      (szFilename As String, szPara As String, szDir As String) As Variant

'���[���p
Public Type MailType
    ServerName        As String
    Clientname        As String
    ClientMailAddress As String
    ClientRealName    As String
    mailPassword      As String
    savefolder        As String
    SendCO            As Integer
    SendName(50)      As String
    JyusinSW          As Integer
End Type
Public MailTabl As MailType

'FTP�T�[�o�p
Public Type FTPsv
    Name As String
    User As String
    Pass As String
End Type

Public mINIfile As String
'Public strData$

'
'##############################################################
Public Sub FTPdataGet(SV As FTPsv, ret As Integer)
'FTP�T�[�o����f�[�^�_�E�����[�h
'    On Local Error GoTo SendFTPerr
    
    Dim fileN() As String, fileC As Long, fileNS As String
    Dim ftpErr  As String

    Dim ftp As BASP21Lib.ftp
    Dim rc As Long
    Dim vv As Variant, vv2 As Variant
    
    Dim i&
    
'    Dim FTPdir$
    
    Set ftp = New BASP21Lib.ftp
    ftp.OpenLog App.Path & "\FTP-log.txt"
    rc = ftp.Connect(SV.Name, SV.User, SV.Pass)
    If rc = 0 Then
        vv = ftp.GetDir("/A0002/appdata") ' �f�B���N�g���ꗗ(�t�@�C�����̂�)
        If IsArray(vv) Then
            fileC = 0
            For Each vv2 In vv
                fileNS = vv2
                fileNS = Trim$(fileNS)
                If UCase$(Right$(fileNS, 4)) = ".DAT" Then
                    fileC = fileC + 1
                    ReDim Preserve fileN(fileC)
                    fileN(fileC) = vv2
                    Debug.Print (fileC), fileN(fileC)
                End If
            Next
        End If
        '�v���f�[�^�̃_�E�����[�h
        For i = 1 To fileC
            rc = ftp.GetFile("/A0002/appdata/" & fileN(i), App.Path & "\AppData")  ' �e�L�X�g�t�@�C���̎�M
            If rc = 1 Then
                rc = ftp.DeleteFile("/A0002/appdata/" & fileN(i))
            End If
        Next i
        
        ftp.Close
    Else
        ftpErr = ftp.GetReply()
    End If
    
    ret = fileC
FTPrw001:
    rc = ftp.CloseLog()
    Set ftp = Nothing
    Erase fileN
Exit Sub

SendFTPerr:
    Set ftp = Nothing
'    Call ErrLog(Now, "FTPrw", (Err.Number & " " & Err.Description))
    Erase fileN
End Sub

Public Sub FTPdataPUT(SV As FTPsv, ret As Integer)
'FTP�T�[�o�փf�[�^�A�b�v���[�h
    On Local Error GoTo SendFTPerr
    
    Dim ftpErr  As String

    Dim ftp As BASP21Lib.ftp
    Dim rc As Long
    
    ret = 0
'    Dim FTPdir$
    
    Set ftp = New BASP21Lib.ftp
    ftp.OpenLog App.Path & "\FTP-log.txt"
    rc = ftp.Connect(SV.Name, SV.User, SV.Pass)
    If rc = 0 Then
        If FileExists(App.Path & "\FTP\master.dat") = True Then
            rc = ftp.PutFile(App.Path & "\FTP\master.dat", "/A0002/DATA", 2) '�t�@�C�����M
            If 1 = rc Then
                Call DelFile(App.Path & "\FTP\" & "MASTER.DAT")
                ret = -1
            End If
        End If
        
        ftp.Close
    Else
        ftpErr = ftp.GetReply()
    End If
    
FTPrw001:
    rc = ftp.CloseLog()
    Set ftp = Nothing
Exit Sub
        
SendFTPerr:
    Set ftp = Nothing
End Sub

Public Sub MailSend(MailTbl As MailType)
   
    Dim i As Integer
    With MailTbl
    
        .SendCO = CInt(GetIni("���[�����M", "���M��", mINIfile))
        For i = 1 To .SendCO
            .SendName(i) = GetIni("���[�����M", "���M��" & CStr(i), mINIfile)
        Next i
    End With
    
    Dim ssb As String
    ssb = (GetIni("���[�����M", "subject", mINIfile))
    
    Dim ret As String
    Dim szServer As String, szTo As String, szFrom As String
    Dim szSubject As String, szBody As String, szFile As String
    
    szServer = MailTbl.ServerName '& ":465:60"    ' SMTP�T�[�o���B�|�[�g�ԍ����w��ł��܂��B
    szServer = "smtp.gmail.com:465:60"               ' SMTP�T�[�o���B�|�[�g�ԍ����w��ł��܂��B"
    szTo = "who@who.com"            ' ���� ' �����̈���ɑ��t����Ƃ��́A�A�h���X���^�u�ŋ�؂��Ă�����ł��w��ł��܂��B
    szTo = MailTbl.SendName(1)
    If MailTbl.SendCO > 1 Then
        szTo = szTo & vbTab & "bcc"
    End If
    For i = 2 To MailTbl.SendCO
        szTo = szTo & vbTab & MailTbl.SendName(i)
    Next i
            '    ' CC���w�肷��ɂ͎��̂悤�ɂ��܂��B
            '        szTo = "who@who.com" & vbTab & "cc" & vbTab & "who2@who2.com" & _
            '           vbTab & "who3@who3.com"
            '    ' BCC���w�肷��ɂ͎��̂悤�ɂ��܂��B
            '        szTo = "who@who.com" & vbTab & "bcc" & vbTab & "who2@who2.com" & _
            '           vbTab & "who3@who3.com"
            '    ' �w�b�_���w�肷��ɂ͎��̂悤�Ƀ^�u�ŋ�؂�A>���w�b�_�̑O��
            '    ' ���܂��B
            '        szTo = "who@who.com" & vbTab & ">Message-ID: 12345"
'    szFrom = MailTbl.ClientMailAddress & vbTab & MailTbl.Clientname & ":" & MailTbl.mailPassword   ' ���M��
    szFrom = MailTbl.ClientMailAddress & vbTab & "a545352322" & ":" & "yo2803ks"   ' ���M��
    szFrom = "<mbkeisya@gmail.com>" & vbTab & "mbkeisya@gmail.com" & ":" & "keisoku2803"   ' ���M��
    szSubject = ssb '     ' ����
    'szBody = "����ɂ��́B" & vbCrLf & "���悤�Ȃ�"   ' �{��' �{�����ŉ��s����ɂ́AvbCrLf���g���܂��B
    szBody = strData '"����f�[�^"
    
    ' �t�@�C����Y�t����Ƃ��́A�t�@�C�������t���p�X�Ŏw�肵�܂��B
    ' �t�@�C���𕡐��w�肷��Ƃ��́A�^�u�ŋ�؂��Ă��������B
    
''    szFile = CurrentDIR & "WaveData.LZH" '& vbTab & "c:\a2.jpeg" ' �t�@�C���Q��
    ' �t�@�C����Y�t���Ȃ��Ƃ��͎��̂悤�ɂ��܂��B
    szFile = ""   ' �t�@�C���Y�t�Ȃ�
    
    ret = SendMail(szServer, szTo, szFrom, szSubject, szBody, szFile)
    
    ' ���M�G���[�̂Ƃ��́A�߂�l�ɃG���[���b�Z�[�W���Ԃ�܂��B
    If Len(ret) <> 0 Then
       'MsgBox "�G���[" & ret
       Call ErrLog(Now, "���[�����M", ret)
    End If
Exit Sub

End Sub

Public Sub MailRead(MailTbl As MailType)
    Dim szServer As String, szUser As String, szPass As String
    Dim szCommand As String, szDir As String
    Dim ar As Variant, v As Variant

    Dim szFilename As String, szPara As String
    Dim retv As Variant
    
    Dim t1 As Date, t2 As Date
    t1 = Now
    Do
        t2 = Now
        If DateDiff("s", DateAdd("s", 2, t1), t2) > 0 Then Exit Do
    Loop
    szServer = MailTbl.ServerName  'SMTP�T�[�o���Ɠ����ł悢�B
                                    '�^�u�ŋ�؂��ă|�[�g�ԍ����w��ł��܂��B
    szUser = MailTbl.Clientname    '���[���A�J�E���g��
    szPass = MailTbl.mailPassword  '�p�X���[�h
    '''      2000/05/20 APOP���T�|�[�g
    '''      APOP �F�؂�����ɂ́A�p�X���[�h�̑O�� "a" �܂��� "A" �� �P��
    '''      �u�����N�����܂��
    '''      "a xxxx" : �T�[�o��APOP ���Ή��Ȃ�ʏ��USER/PASS ���������܂��B
    '''      "A xxxx" : �T�[�o��APOP ���Ή��Ȃ�G���[�ɂȂ�܂��B
       
#If DebugVersion Then
    szCommand = "SAVEALL"  '�R�}���h�@���[���̂P���ڂ���R���ڂ܂ł���M
#Else
    szCommand = "SAVEALLD"  '�R�}���h�@���[���̂P���ڂ���R���ڂ܂ł���M
#End If
    
    szDir = App.Path & "\MailData" 'MailTabl.savefolder '��M�������[����ۑ�����f�B���N�g��
    
    ar = RcvMail(szServer, szUser, szPass, szCommand, szDir)
    
'    Dim smCMD$
    '�߂�l���Ԃ�ϐ��́AVariant�^�C�v���w�肷�邱�ƁB
    '��M�������[���P�ʂ��ƂɃt�@�C�����쐬����܂��B
    '���[���ɓY�t���ꂽ�t�@�C���́A�{���Ƌ��ɂP�̃t�@�C���Ɋ܂܂�܂��B
    'ReadMail�֐��œY�t�t�@�C������o���܂��B
    If IsArray(ar) Then   '����I������SAVE�R�}���h�̖߂�l�́A�z��ɂȂ�܂��B
        For Each v In ar
            'Debug.Print v     '���[���f�[�^���ۑ����ꂽ�t�@�C�������t���p�X�Ŗ߂�܂��B
                              '���̃t�@�C������ReadMail�̃p�����[�^�Ƃ��ēn���܂��B
            szFilename = v  ' �t�@�C�����ɂ�RcvMail�̖߂�l�̔z�񂩂�t�@�C������ݒ�
            szPara = "subject:from:date:"  ' �w�b�_�[�̎w��
                                           ' nofile: �Ƃ���ƓY�t�t�@�C����ۑ����܂���B
            retv = ReadMail(szFilename, szPara, szDir)
            If IsArray(retv) Then
'                If 0 < InStr(retv(0), "�C���^�[�o���ύX") Then
'                    smCMD = retv(3)
'                    Call gCMD(smCMD)
'                ElseIf 0 < InStr(retv(0), "�x�񔭐�") Then
'                    smCMD = retv(3)
'                    Call KeihouARM(smCMD)
'                End If
                
                'For Each v2 In retv
                '     Debug.Print v2
                'Next
            Else
                'Debug.Print retv
            End If
        Next
    Else
        'Debug.Print ar      '�G���[�������́A�z��łȂ����b�Z�[�W���߂�܂��B
    End If

End Sub

