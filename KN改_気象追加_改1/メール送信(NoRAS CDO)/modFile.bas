Attribute VB_Name = "modFile"
Option Explicit

'
'[実在のパスが存在するか調べる]
'
'■ 引数
'strPathName:フルパスのファイル名
'□ 戻り値:パスの有無(True=存在する ,False=存在しない)
Public Function FilePasExists(ByVal strPathName As String) As Boolean
  Dim strResult As String
  On Error Resume Next
  If strPathName = "" Then Exit Function
  'フォルダーに \ をつけるかどうか識別
  If Right(strPathName, 1) <> "\" Then strPathName = strPathName & "\"

  strResult = Dir(strPathName & "*.*", vbDirectory)
  FilePasExists = IIf(strResult = "", False, True)

  Err = 0
End Function
'
'[ファイルの有無を調査する]
'
'■ 引数
'FileName:フルパスのファイル名
'□ 戻り値:パスの有無(True=存在する ,False=存在しない)
Public Function FileExists(ByVal FILENAME As String) As Boolean
  Dim TempAttr As Integer

  If (Len(FILENAME) = 0) Or (InStr(FILENAME, "*") > 0) Or _
                                                 (InStr(FILENAME, "?") > 0) Then
     FileExists = False
     Exit Function
  End If
  On Error GoTo ErrorFileExist
  ' ファイルの属性を得る
  TempAttr = GetAttr(FILENAME)
  ' ディレクトリであるかどうか調べる
  FileExists = ((TempAttr And vbDirectory) = 0)
  GoTo ExitFileExist
ErrorFileExist:
  FileExists = False
  Resume ExitFileExist
ExitFileExist:
  On Error GoTo 0
End Function


