Attribute VB_Name = "WinAPI"
Option Explicit

'INIファイルより読み込み
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'INIファイルに書き込み
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'
'※ 注意！！    情報を格納するバッファは必ず長さを指定すること。
'               WritePrivateProfileStringでファイルがなかった場合は作成される。
'
'<<解説>>GetPrivateProfileString( セクション as String
'                                 キー as String
'                                 デフォルト as String          キーがなかった場合の値
'                                 情報 as Any                   情報を格納するバッファ
'                                 情報の長さ as Long            バッファの長さ
'                                 INIファイルパス as String
'
'<<解説>>WritePrivateProfileString( セクション as String
'                                   キー as Any
'                                   書き込む情報 as String
'                                   INIファイルパス as String
'

'**********************************
'        INIファイル読み込み
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
'        INIファイル書き込み
'**********************************
Public Function WriteIni(section As String, key As String, str As String, INIFile As String) As Long
    Dim ret As Long
    ret = WritePrivateProfileString(section, key, str, INIFile)
    WriteIni = ret
End Function
