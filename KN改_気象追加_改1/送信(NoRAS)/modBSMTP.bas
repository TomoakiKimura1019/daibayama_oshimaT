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

