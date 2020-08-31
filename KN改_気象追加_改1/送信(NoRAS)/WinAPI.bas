Attribute VB_Name = "WinAPI"
Option Explicit

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
