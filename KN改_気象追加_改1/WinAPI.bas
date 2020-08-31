Attribute VB_Name = "WinAPI"
Option Explicit

'�l�b�g���[�N�h���C�u��ChDrive����
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias _
                           "SetCurrentDirectoryA" (ByVal CurrentDir As String) As Long

'INI�t�@�C�����ǂݍ���
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'INI�t�@�C���ɏ�������
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Public Const SW_SHOW = 5

'�f�X�N�g�b�v�E�B���h�E�̃n���h�����擾����API
Public Declare Function GetDesktopWindow Lib "user32" () As Long

'�w�肳�ꂽ�t�@�C�����I�[�v���A���邢�͕\������API
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
'
'�� ���ӁI�I    �����i�[����o�b�t�@�͕K���������w�肷�邱�ƁB
'               WritePrivateProfileString�Ńt�@�C�����Ȃ������ꍇ�͍쐬�����B
'
'<<���>>GetPrivateProfileString( �Z�N�V���� as String
'                                 �L�[ as String
'                                 �f�t�H���g as String          �L�[���Ȃ������ꍇ�̒l
'                                 ��� as Any                   �����i�[����o�b�t�@
'                                 ���̒��� as Long            �o�b�t�@�̒���
'                                 INI�t�@�C���p�X as String
'
'<<���>>WritePrivateProfileString( �Z�N�V���� as String
'                                   �L�[ as Any
'                                   �������ޏ�� as String
'                                   INI�t�@�C���p�X as String
'

'**********************************
'        INI�t�@�C���ǂݍ���
'**********************************
Public Function GetIni(section As String, key As String, INIFile As String) As String
    Dim StrBuf As String * 1024
    Dim ret As Long
    Dim EBuf As String
    Dim i As Long
    
    StrBuf = String$(1024, Chr$(0))
    ret = GetPrivateProfileString(section, key, "", StrBuf, 1024, INIFile)
    
    i = InStr(StrBuf, vbNullChar)
    If i <> 0 Then
        EBuf = Left$(StrBuf, i - 1)
    Else
        EBuf = StrBuf
    End If
    
    GetIni = EBuf

End Function

'**********************************
'        INI�t�@�C����������
'**********************************
Public Function WriteIni(section As String, key As String, str As String, INIFile As String) As Long
    Dim ret As Long
    ret = WritePrivateProfileString(section, key, str, INIFile)
    WriteIni = ret
End Function

