Attribute VB_Name = "Module1"
Option Explicit

' プロセス一覧ＡＰＩ
Public Const TH32CS_SNAPHEAPLIST = 1
Public Const TH32CS_SNAPPROCESS = 2
Public Const TH32CS_SNAPTHREAD = 4
Public Const TH32CS_SNAPMODULE = 8
Public Const TH32CS_SNAPALL = 15
Public Const SIZEOF_PROCESSENTRY32 As Long = 296
Public Type PROCESSENTRY32
    Size As Long
    RefCount As Long
    ProcessID As Long
    HeapID As Long
    ModuleID As Long
    ThreadCount As Long
    ParentProcessID As Long
    BasePriority As Long
    Flags As Long
    FileName As String * 260
End Type
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
    ByVal Flags As Long, ByVal ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" ( _
    ByVal lngHandleshot As Long, _
    ByRef ProcessEntry As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" ( _
    ByVal lngHandle As Long, _
    ByRef ProcessEntry As PROCESSENTRY32) As Long
' プロセスの起動／終了ＡＰＩ
Public Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long
Public Const SYNCHRONIZE       As Long = &H100000
Public Const PROCESS_TERMINATE As Long = &H1

Public Function GetFullPasToFileName(ByVal FullPas As String) As String
  Dim i As Integer, tmp As String

  For i = Len(FullPas) To 1 Step -1
       Select Case Mid$(FullPas, i, 1)
           Case "\", ":"
                  GetFullPasToFileName = Mid$(FullPas, i + 1)
                  Exit For
       End Select
  Next i
End Function

'■ 引数
'strFileName:フルパスのファイル名
'□ 戻り値:ファイル名
Public Function GetPathNameToFullPas(ByVal strFileName As String) As String
Dim intPos As Integer
Dim strPathOnly As String
Dim intLoopCount As Integer
On Error Resume Next
  Err = 0
  intPos = Len(strFileName)
  'すべての '/' 記号を '\'記号に変更します。
     For intLoopCount = 1 To Len(strFileName)
           If Mid(strFileName, intLoopCount, 1) = "/" Then
           Mid(strFileName, intLoopCount, 1) = "\"
           End If
     Next intLoopCount

     If InStr(strFileName, "\") = intPos Then
         If intPos > 1 Then intPos = intPos - 1
     Else
         Do While intPos > 0
               If Mid(strFileName, intPos, 1) <> "\" Then
                   intPos = intPos - 1
               Else
                   Exit Do
               End If
          Loop
     End If

     If intPos > 0 Then
      strPathOnly = Left(strFileName, intPos)
      If Right(strPathOnly, 1) = ":" Then strPathOnly = strPathOnly & "\"
     Else
     strPathOnly = CurDir
     End If

     If Right(strPathOnly, 1) = "\" Then
       strPathOnly = Left(strPathOnly, Len(strPathOnly) - 1)
     End If

     GetPathNameToFullPas = strPathOnly
     Err = 0
End Function


