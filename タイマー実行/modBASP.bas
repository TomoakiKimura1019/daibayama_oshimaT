Attribute VB_Name = "modBASP"
Option Explicit

Public Sub FindDataFile(ByVal id%, ByVal fdir As String, ByVal id2 As Integer)
'�t�@�C������
'�t�@�C������������A�z��Ƀp�X�����擾
'fDir : �����f�B���N�g��

    Dim FileList() As String
    Dim i As Long

    Dim ret As String

    Dim tFilename() As String
    Dim aIndex As Long
    aIndex = -1

        If GetTargetFiles(FileList, fdir, "csv") Then
        '�t�@�C������z��Ɏ擾
        For i = 0 To UBound(FileList)
'            Debug.Print FileList(i)
            '����̌^���̃t�@�C����I��
            'ret = Match("/\d{1,4}_\d{1,2}_BV\d{1}-[XY]_disp.txt/", FindFileName(FileList(i)))
'            ret = Match("/\d{1,4}_\d{1,2}_strain.txt/", FindFileName(FileList(i)))
'            If ret = "1" Then
                aIndex = aIndex + 1
                ReDim Preserve tFilename(aIndex) As String
                tFilename(aIndex) = FindFileName(FileList(i))
'            End If
        Next i
        '���������t�@�C�������\�[�g
        If -1 < aIndex Then
            s_ShellSort tFilename(), (aIndex)
        End If

        If aIndex = -1 Then Exit Sub
        
        'frmTDSdataget.StatusBar1.Panels(1).Text = "found"
        
        'FTP�ő��M
        Dim rc As Integer
'        If id = 1 Then
''            Call SendFTP(rc, fdir, tFilename(), TdsDataPath) ' �������ԈႢ
'            Call SendFTP(rc, fdir, tFilename(), TDSFTPpath) '
'        End If
        If id = 2 Then
            'frmTDSdataget.StatusBar1.Panels(1).Text = "SendFTP start"
            Call SendFTP(rc, fdir, tFilename(), "/array1/share2/���L����/�v�������̏���/etc")
        End If
        
    End If

End Sub

Public Function FTPpathname(ByVal tFilename As String, sYY$, sMM$, sDD$) As String
'�t�@�C��������ړI��FTP�f�B���N�g�����𐶐�

'    Dim sYY As String
'    Dim sMM As String
'    Dim sDD As String
    Dim sNN As String
    
    '2009-10-12_10-00.dat
    sYY = Mid$(tFilename, 1, 4)
    sMM = Mid$(tFilename, 6, 2)
    sDD = Mid$(tFilename, 9, 2)
    sNN = "/" & sYY & "/" & sMM & "/" & sDD
    
    FTPpathname = sNN
    
End Function

Public Sub SendFTP(ret As Integer, fdir As String, fPath() As String, FTPpath$)
Dim SendPath, rSettei
'�����[�g�փt�@�C���𑗐M���܂�������t�@�C���̑��M���ł��܂��
'
'rc = ftp.PutFile(local,remote[,type])
'  local [in]  : ���M����t�@�C�������t���p�X�Ŏw��B
'                �����t�@�C���̎w��́A "a*.txt" �A"*"�A"*.html" �Ȃǂ̂悤�� "*" ���g���B
'                ��F c:\html\a.html --- html�f�B���N�g����a.html
'                     c:\html\*.html --- html�f�B���N�g���� .html �t�@�C�����ׂ�
'                     c:\html\*      --- html�f�B���N�g���̂��ׂẴt�@�C��
' remote [in]  : �����[�g�̃f�B���N�g�����B"" �́A�J�����g�f�B���N�g���B
' Enum in: ���M����f�[�^�`�������̂悤�Ɏw��
'  0 : ASCII�i�ȗ��l)�Btxt/html �Ȃǂ̃e�L�X�g�t�@�C���̏ꍇ�B
'  1 : �o�C�i���Bjpg/gif/exe/lzh/tar.gz �Ȃǂ̃o�C�i���t�@�C���̏ꍇ�B
'  2 : ASCII + �ǉ�(Append)���[�h�B
'  3 : �o�C�i�� + �ǉ�(Append)���[�h�B
'
'  rc [out]: ���ʃR�[�h�������ŕԂ���܂��
'  1 �ȏ�:   ����I������M�����t�@�C�����
'  0     :   �Y������t�@�C���Ȃ��
'  1,0�ȊO : �G���[�BGetReply���\�b�h��FTP�����e�L�X�g�ŏڍׂ𒲂ׂĂ��������B
'��:
'rc = ftp.PutFile("c:\html\index.html", "html")  ' �e�L�X�g�t�@�C���̑��M
'rc = ftp.PutFile("c:\html\*.html", "html")      ' �e�L�X�g�t�@�C���̑��M
'rc = ftp.PutFile("c:\html\*.html", "html", 2)     ' �e�L�X�g�t�@�C����Append���[�h���M
'rc = ftp.PutFile("c:\html\images\*", "html/images", 1) ' �o�C�i���t�@�C���̑��M
    
    Dim ServerIP As String
    Dim i As Integer
    Dim tFile As String
    
    Dim sYY As String
    Dim sMM As String
    Dim sDD As String
    Dim fpSW As Boolean
    ret = 0
    On Local Error GoTo SendFTPerr
    
    Dim ftpErr  As String
    Dim rc As Long
    Dim vv As Variant, vv2 As Variant
''    Dim ftp As Object
''    Set ftp = CreateObject("basp21.FTP")
    Dim ftp As BASP21Lib.ftp
    Set ftp = New BASP21Lib.ftp
    
    ftp.OpenLog App.Path & "\FTP-log.txt"
'    rc = ftp.Connect("172.16.60.219", "a-tic", "keisoku")  '�{��
If Command$ = "TEST" Then
    ServerIP = "172.16.60.99"
    rc = ftp.Connect(ServerIP, "anonymous", "")  '
Else
    ServerIP = "180.43.16.132"
'    ServerIP = "172.16.65.96"
    rc = ftp.Connect(ServerIP & ":49621", "otonaka", "atic2803")  '�{��
End If
'    rc = ftp.Connect("60.43.239.36", "chikah", "zbeba+nn")  '�O�̖{��
    If rc = 0 Then
        'frmTDSdataget.StatusBar1.Panels(1).Text = "FTP connect"
        ' passive���[�h�ɂ���
        ftp.Command ("PASV") ' ��x�ďo���� OK
        '�v���f�[�^�̃A�b�v���[�h
        For i = 0 To UBound(fPath)
            tFile = FTPpathname(fPath(i), sYY, sMM, sDD)
            If Left$(tFile, 1) = "/" Then tFile = Right$(tFile, Len(tFile) - 1)
            rc = ftp.Command("CWD " & FTPpath & "/" & sYY)   '�f�B���N�g���ړ�
            If Not (rc = 2) Then
                ftpErr = ftp.GetReply()
                'If InStr(ftpErr, "No such file or directory") > 0 Then
                If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY)    '�f�B���N�g���쐬
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)    '�f�B���N�g���쐬
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���쐬
                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���ړ�
                End If
            Else
                rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM)    '�f�B���N�g���ړ�
                If Not (rc = 2) Then
                    ftpErr = ftp.GetReply()
                    'If InStr(ftpErr, "No such file or directory") > 0 Then
                    If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)    '�f�B���N�g���쐬
                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���쐬
                        rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���ړ�
                    End If
                Else
                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���ړ�
                    If Not (rc = 2) Then
                        ftpErr = ftp.GetReply()
                        'If InStr(ftpErr, "No such file or directory") > 0 Then
                        If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                            rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���쐬
                            rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    '�f�B���N�g���ړ�
                        End If
                    End If
                End If
            End If
            
            rc = ftp.PutFile(fdir & "\" & fPath(i), "", 1) '�t�@�C�����M
            
            If rc = 1 Then
                fpSW = False
                vv = ftp.GetDir("") ' �f�B���N�g���ꗗ(�t�@�C����)
                If IsArray(vv) Then
                    For Each vv2 In vv
                        If vv2 = fPath(i) Then
                            fpSW = True
                            Exit For
                        End If
                    Next
                End If
                If fpSW = True Then
                    sFileDelete fdir & "\" & fPath(i)
                End If
            End If
        Next i
        ftp.Close
        'frmTDSdataget.StatusBar1.Panels(1).Text = ""
    Else
        ftpErr = ftp.GetReply()
        'frmTDSdataget.StatusBar1.Panels(1).Text = "FTP connect error"
    End If
    rc = ftp.CloseLog()
    
    Set ftp = Nothing
    ret = -1
Exit Sub
SendFTPerr:
    Set ftp = Nothing
End Sub

