Attribute VB_Name = "RASmodule"
Option Explicit

Public Type RAS1
    eName As String
    User  As String
    Pw    As String
    iEntNo As Integer
End Type
Public RAS(2) As RAS1

Public m_Rasmon
Public iEntNo As Integer
Public strData As String
Public ConnectCK  As Integer

'Sub RASinit()
'    With RAS
'        .eName = "mopera"
'        .User = ""
'        .Pw = ""
'    End With
'
'RasInitial
'End Sub


