Attribute VB_Name = "modFile"
Option Explicit

'�t�@�C���n���h�����擾����
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

'�t�@�C������ǂݍ���
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

'�t�@�C���n���h�������
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const OPEN_EXISTING = 3

'�t�@�C���^�C��
Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
     dwFileAttributes As Long       '�t�@�C������
     ftCreationTime As FILETIME     '�쐬��
     ftLastAccessTime As FILETIME   '�A�N�Z�X��
     ftLastWriteTime As FILETIME    '�X�V��
     nFileSizeHigh As Long          '�t�@�C���T�C�Y(Byte)
     nFileSizeLow As Long           '�t�@�C���T�C�Y(Byte)
     dwReserved0 As Long            '���g�p
     dwReserved1 As Long            '���g�p
     cFileName As String * 260      '�t�@�C����
     cAlternate As String * 14      '�t�@�C�����t�H�[�}�b�g��
End Type
'�t�@�C���̌������J�n����
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'�t�@�C���̌����𑱍s����
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'�����n���h�������
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Type DIR_FILE_LIST
    FILENAME As String
    IsDirectory As Boolean
End Type

'�p�X����p
Private Declare Function PathFindFileName Lib "SHLWAPI.DLL" Alias "PathFindFileNameA" _
                                (ByVal pszPath As String) As Long
Private Const MAX_PATH = 260
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                (pDest As Any, _
                                 pSource As Any, _
                                 ByVal ByteLen As Long)

Private Declare Function PathRemoveBackslash Lib "SHLWAPI.DLL" Alias "PathRemoveBackslashA" _
                                (ByVal pszPath As String) As Long

Private Declare Function PathRemoveFileSpec Lib "SHLWAPI.DLL" Alias "PathRemoveFileSpecA" _
                                (ByVal pszPath As String) As Long


' �t�@�C������p��API-----------------------------------------------------
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3&              ' �폜
Private Const FO_COPY = &H2                 ' �R�s�[
Private Const FO_MOVE = &H1                 ' �ړ�
Private Const FO_RENAME = &H4               ' �t�@�C�����ύX
Private Const FOF_ALLOWUNDO = &H40&         ' ���ݔ���
Private Const FOF_NOCONFIRMATION = &H10&    ' �m�F�_�C�A���O��\�����Ȃ�
Private Const FOF_NOERRORUI = &H400&        ' �G���[�_�C�A���O��\�����Ȃ�
Private Const FOF_MULTIDESTFILES = &H1&     ' �����t�@�C�����w�肷��

'Private Const FO_MOVE As Long = &H1
'Private Const FO_COPY As Long = &H2
'Private Const FO_DELETE As Long = &H3
'Private Const FO_RENAME As Long = &H4
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
'Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
'Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
'Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200
'Private Const FOF_NOERRORUI As Long = &H400

' �t�@�C������������B
' RootPath          : �������J�n�����̃f�B���N�g��
' InputPathName     : ��������t�@�C����
' OutputPathBuffer  : ���������t�@�C�������i�[����o�b�t�@�B
' �߂�l            : �������0�ȊO��Ԃ��B
Private Declare Function SearchTreeForFile Lib "imagehlp.dll" _
    (ByVal RootPath As String, _
     ByVal InputPathName As String, _
     ByVal OutputPathBuffer As String) As Long

Private Const MAX_PATHp = 512
Private Const MAX_PATH_PLUS1 = MAX_PATHp + 1

Dim ttt As String

Public Function CheckDataFile(ByVal fdir As String) As Long
'�t�@�C������
'�t�@�C������������A�z��Ƀp�X�����擾
'fDir : �����f�B���N�g��

    Dim FileList() As String
    Dim i As Long


    Dim tFilename() As String
    Dim aIndex As Long
    aIndex = -1

    If GetTargetFiles(FileList, fdir, "dat") Then
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

    '    If aIndex = -1 Then Exit Function
    End If
    CheckDataFile = aIndex + 1
End Function

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


'-------------------------------------------------------------------
' �֐��� �F ReadFileUsingAPIFunc
' �@�\ �F �t�@�C�����當�����ǂݍ��� �e�L�X�g�^�ϐ��ɑ������
' ���� �F �~(in) srcText �c �������\������e�L�X�g�{�b�N�X
'           (in) fPath �c �ǂݍ��ރt�@�C���̃p�X
' �Ԃ�l �F  true : �ǂݎ�萬��
'           False : �ǂ߂Ȃ�����
'-------------------------------------------------------------------
Private Function ReadFileUsingAPIFunc(ByVal Fpath As String) As Boolean

    Dim hFile As Long       '�t�@�C���̃n���h��
    Dim FileSize As Long    '�t�@�C���T�C�Y
    Dim gBinData() As Byte  '�擾�f�[�^
    Dim outFileSize As Long

    '�t�@�C�����J��(READ)
    hFile = CreateFile(Fpath, GENERIC_READ, 0&, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile = -1 Then
        ReadFileUsingAPIFunc = False
        Exit Function
    Else
    End If

    '�t�@�C���T�C�Y�擾
    FileSize = FileLen(Fpath)
    If FileSize < 567 Then
        '�t�@�C�������
        Call CloseHandle(hFile)
        ReadFileUsingAPIFunc = False
    End If

    '�ϐ�������
    ReDim Preserve gBinData(FileSize - 1) As Byte

    '�t�@�C���ǂݍ���
    Call ReadFile(hFile, gBinData(0), FileSize, outFileSize, 0&)

    'ANSI �� Unicode�ϊ�
'    srcText.Text = StrConv(gBinData(), vbUnicode)
    ttt = StrConv(gBinData(), vbUnicode)

    '�t�@�C�������
    Call CloseHandle(hFile)
    
    ReadFileUsingAPIFunc = True

End Function

'-----------------------------------------------------------------------
' �֐��� �F GetTargetFiles
' �@�\   �F �f�B���N�g���ȉ��̎w��g���q�̃t�@�C�����擾����
' ����   �F (in/out) Files �c �擾�����t�@�C�����i�[����z��
'           (in)DirName    �c �������f�B���N�g��
'           (in)Extension  �c �g���q
' �Ԃ�l �F True�F�������f�B���N�g���͑��݂���  False�F���݂��Ȃ�
'-----------------------------------------------------------------------
Public Function GetTargetFiles(ByRef Files() As String, _
                                ByVal DirName As String, _
                                ByVal Extension As String) As Boolean

    Dim udtWin32 As WIN32_FIND_DATA
    Dim hFile As Long
    Dim ArrayIndex As Long
    Dim FileListNum As Long
    Dim i As Long
    Dim UdtDFL() As DIR_FILE_LIST

    '�Ō���� \ ���t���Ă���ꍇ�͎��
    If Right$(DirName, 1) = "\" Then DirName = Left$(DirName, Len(DirName) - 1)

    '�����J�n
    hFile = FindFirstFile(DirName, udtWin32)
    '�t�@�C���E�f�B���N�g�������݂��Ȃ��ꍇ�͏I��
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Call FindClose(hFile)

    '�f�B���N�g���ȉ��̃t�@�C���E�f�B���N�g�����擾����
    FileListNum = GetFileList(UdtDFL, DirName)
    If FileListNum = (-1) Then Exit Function

    For i = 0 To FileListNum
        '�f�B���N�g���ł���
        If UdtDFL(i).IsDirectory Then
            Call GetTargetFiles(Files, DirName & "\" & UdtDFL(i).FILENAME, Extension)
        '�t�@�C���ł���
        Else
            '�t�@�C���̊g���q���w��g���q�Ɠ�����
            If UCase$(Right$(UdtDFL(i).FILENAME, Len(Extension))) = UCase$(Extension) Then
                '������s�� Files �͔z�񖳂��Ȃ̂�UBound()�ŃG���[�ƂȂ�
                '�����������邽�߂̋����G���[�������W�b�N
                On Error Resume Next
                ArrayIndex = UBound(Files) + 1
                On Error GoTo 0

                '�������[�m��
                ReDim Preserve Files(ArrayIndex) As String

                '�t���p�X�Ńt�@�C�������i�[
                Files(ArrayIndex) = DirName & "\" & UdtDFL(i).FILENAME
            End If
        End If
    Next i

    GetTargetFiles = True

End Function

'-----------------------------------------------------------------------
' �֐��� �F GetFileList
' �@�\   �F �f�B���N�g���̃t�@�C�����擾����
' ����   �F (in/out) UdtDFL �c DIR_FILE_LIST�\���̂̔z��
'           (in)DirName     �c �������f�B���N�g��
' �Ԃ�l �F UdtDFL�z��   �t�@�C�������݂��Ȃ��ꍇ�F-1
'-----------------------------------------------------------------------
Private Function GetFileList(ByRef UdtDFL() As DIR_FILE_LIST, _
                            ByVal DirName As String) As Long

    Dim udtWin32 As WIN32_FIND_DATA
    Dim hFile As Long
    Dim ArrayIndex As Long
    Dim Win32FileName As String

    ArrayIndex = (-1)

    '�����J�n
    hFile = FindFirstFile(DirName & "\*", udtWin32)
    Do
        '���X�A�ĕ`��
        If ArrayIndex Mod 10 = 0 Then DoEvents

        '�t�@�C�����擾
        Win32FileName = Left$(udtWin32.cFileName, _
                              InStr(udtWin32.cFileName, Chr$(0)) - 1)

        '�e�f�B���N�g���A�J�����g�f�B���N�g���łȂ�
        If Left$(Win32FileName, 1) <> "." Then
            ArrayIndex = ArrayIndex + 1
            ReDim Preserve UdtDFL(ArrayIndex) As DIR_FILE_LIST
            '�t�@�C�����A�t�@�C���������擾
            With UdtDFL(ArrayIndex)
                .FILENAME = Win32FileName
                .IsDirectory = CBool(udtWin32.dwFileAttributes And vbDirectory)
            End With
        End If
    Loop While FindNextFile(hFile, udtWin32) <> 0

    Call FindClose(hFile)

    GetFileList = ArrayIndex

End Function

Public Sub s_ShellSort(ByRef sArray() As String, ByVal Num As Integer)
   Dim Span As Integer
   Dim i As Integer
   Dim j As Integer
   Dim TMP As String
   
   Span = Num \ 2
   Do While Span > 0
      For i = Span To Num - 1
         j% = i% - Span + 1
         For j = (i - Span + 1) To 0 Step -Span
            If sArray(j) <= sArray(j + Span) Then Exit For
            ' ���Ԃ̈قȂ�z��v�f�����ւ��܂�.
            TMP = sArray(j)
            sArray(j) = sArray(j + Span)
            sArray(j + Span) = TMP
         Next j
      Next i
      Span = Span \ 2
   Loop
End Sub
'
' �t�@�C���������o���B
'
Public Function FindFileName(ByVal strFileName As String) As String
    ' strFileName : �t���p�X�̃t�@�C����
    ' �߂�l      : �t�@�C�����������Ԃ�B
    Dim strBuffer   As String
    Dim lngResult   As Long
    Dim bytStr()    As Byte

    lngResult = PathFindFileName(strFileName)
    If lngResult <> 0 Then
        ' (MAX_PATH + 1)�̃o�C�g�z���p�ӂ���B
        ReDim bytStr(MAX_PATH + 1) As Byte
        ' �m�ۂ����o�C�g�z��ɓ���ꂽ�ʒu�̃��������R�s�[����B
        MoveMemory bytStr(0), ByVal lngResult, MAX_PATH + 1
        ' �z��𕶎���ɕϊ�����B
        strBuffer = StrConv(bytStr(), vbUnicode)
        ' NULL�����܂ł�؂�o���B
        FindFileName = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If
End Function

'
' �p�X���������o���B
'
Public Function RemoveFileSpec(ByVal strPath As String) As String
    ' strPath : �t���p�X�̃t�@�C����
    ' �߂�l  : �p�X��

    Dim lngResult As Long
    lngResult = PathRemoveFileSpec(strPath)
    If lngResult <> 0 Then
        If InStr(strPath, vbNullChar) > 0 Then
            RemoveFileSpec = Left$(strPath, InStr(strPath, vbNullChar) - 1)
        Else
            RemoveFileSpec = strPath
        End If
    End If
End Function

Public Sub sFileDelete(DelFile As String)
    '**************************************************************
    '* SHFileOperation�֐����Ăяo���t�@�C�������ݔ��ɑ���@�@�@�@*
    '* meForm�@ = �_�C�A���O��\������Form�@�@�@�@�@�@�@�@�@�@�@�@*
    '* DelFile�@= �폜����t�@�C�����iPath�t�j�@�@�@�@�@�@�@�@�@�@*
    '*�@�@�@�@�@�@�����̃t�@�C�����w�肷��ꍇvbNullChar�ŋ�؂�@*
    '*�@�@�@�@�@�@�Ō�͓��vbNullChar�ŏI���@�@�@�@�@�@�@�@�@*
    '**************************************************************
    On Error Resume Next
    Dim lpFileOp As SHFILEOPSTRUCT
    Dim Result   As Long
    Dim MyFlag   As Long

'�S�~���̏ꍇ
    '�w����@�͂��D�݂Őݒ肵�ĉ������B
    MyFlag = FOF_ALLOWUNDO                  '���ݔ���
    MyFlag = MyFlag + FOF_NOCONFIRMATION    '�m�F���Ȃ�
    ''MyFlag = MyFlag + FOF_MULTIDESTFILES    '�����t�@�C��
    MyFlag = MyFlag + FOF_NOERRORUI         '�G���[�̃_�C�A���O���\��

    ' �t�@�C������Ɋւ�������w��
    With lpFileOp
        .hWnd = App.hInstance ' .hWnd       ' �_�C�A���O�̐e�E�B���h�E�n���h�����w��
        .wFunc = FO_DELETE       ' �폜���w��
        .pFrom = DelFile         ' �폜����f�B���N�g�����w��
       ' .pTo = �����̃t�@�C�����E�f�B���N�g����
        .fFlags = MyFlag         '������@���w��
    End With

    ' �t�@�C����������s
    Result = SHFileOperation(lpFileOp)

    On Error GoTo 0

End Sub

Public Sub sFileMove(DelFile As String)
    '**************************************************************
    '* SHFileOperation�֐����Ăяo���t�@�C�������ݔ��ɑ���@�@�@�@*
    '* meForm�@ = �_�C�A���O��\������Form�@�@�@�@�@�@�@�@�@�@�@�@*
    '* DelFile�@= �폜����t�@�C�����iPath�t�j�@�@�@�@�@�@�@�@�@�@*
    '*�@�@�@�@�@�@�����̃t�@�C�����w�肷��ꍇvbNullChar�ŋ�؂�@*
    '*�@�@�@�@�@�@�Ō�͓��vbNullChar�ŏI���@�@�@�@�@�@�@�@�@*
    '**************************************************************
    On Error Resume Next
    Dim lpFileOp As SHFILEOPSTRUCT
    Dim Result   As Long
    Dim MyFlag   As Long

'�S�~���̏ꍇ
'    '�w����@�͂��D�݂Őݒ肵�ĉ������B
'    MyFlag = FOF_ALLOWUNDO                  '���ݔ���
'    MyFlag = MyFlag + FOF_NOCONFIRMATION    '�m�F���Ȃ�
'    ''MyFlag = MyFlag + FOF_MULTIDESTFILES    '�����t�@�C��
'    MyFlag = MyFlag + FOF_NOERRORUI         '�G���[�̃_�C�A���O���\��
'
'    ' �t�@�C������Ɋւ�������w��
'    With lpFileOp
'        .hWnd = App.hInstance ' .hWnd       ' �_�C�A���O�̐e�E�B���h�E�n���h�����w��
'        .wFunc = FO_DELETE       ' �폜���w��
'        .pFrom = DelFile         ' �폜����f�B���N�g�����w��
'       ' .pTo = �����̃t�@�C�����E�f�B���N�g����
'        .fFlags = MyFlag         '������@���w��
'    End With

    MyFlag = FOF_NOCONFIRMMKDIR                  '���ݔ���
    MyFlag = MyFlag + FOF_NOCONFIRMATION    '�m�F���Ȃ�
    ''MyFlag = MyFlag + FOF_MULTIDESTFILES    '�����t�@�C��
    MyFlag = MyFlag + FOF_NOERRORUI         '�G���[�̃_�C�A���O���\��

    ' �t�@�C������Ɋւ�������w��
    With lpFileOp
        .hWnd = App.hInstance ' .hWnd       ' �_�C�A���O�̐e�E�B���h�E�n���h�����w��
        .wFunc = FO_MOVE       ' �폜���w��
        .pFrom = DelFile         ' �폜����f�B���N�g�����w��
        .pTo = App.Path & "\tmp\" '�����̃t�@�C�����E�f�B���N�g����
        .fFlags = MyFlag         '������@���w��
    End With

    ' �t�@�C����������s
    Result = SHFileOperation(lpFileOp)

    On Error GoTo 0

End Sub

