Attribute VB_Name = "modFile"
Option Explicit

'
'[���݂̃p�X�����݂��邩���ׂ�]
'
'�� ����
'strPathName:�t���p�X�̃t�@�C����
'�� �߂�l:�p�X�̗L��(True=���݂��� ,False=���݂��Ȃ�)
Public Function FilePasExists(ByVal strPathName As String) As Boolean
  Dim strResult As String
  On Error Resume Next
  If strPathName = "" Then Exit Function
  '�t�H���_�[�� \ �����邩�ǂ�������
  If Right(strPathName, 1) <> "\" Then strPathName = strPathName & "\"

  strResult = Dir(strPathName & "*.*", vbDirectory)
  FilePasExists = IIf(strResult = "", False, True)

  Err = 0
End Function
'
'[�t�@�C���̗L���𒲍�����]
'
'�� ����
'FileName:�t���p�X�̃t�@�C����
'�� �߂�l:�p�X�̗L��(True=���݂��� ,False=���݂��Ȃ�)
Public Function FileExists(ByVal FILENAME As String) As Boolean
  Dim TempAttr As Integer

  If (Len(FILENAME) = 0) Or (InStr(FILENAME, "*") > 0) Or _
                                                 (InStr(FILENAME, "?") > 0) Then
     FileExists = False
     Exit Function
  End If
  On Error GoTo ErrorFileExist
  ' �t�@�C���̑����𓾂�
  TempAttr = GetAttr(FILENAME)
  ' �f�B���N�g���ł��邩�ǂ������ׂ�
  FileExists = ((TempAttr And vbDirectory) = 0)
  GoTo ExitFileExist
ErrorFileExist:
  FileExists = False
  Resume ExitFileExist
ExitFileExist:
  On Error GoTo 0
End Function


