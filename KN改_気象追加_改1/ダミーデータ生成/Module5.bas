Attribute VB_Name = "Module5"
Option Explicit

Public Type sTDS
    CH As String
    DAT As String
End Type

Public Type ct
  id         As Integer
  Field      As Integer
  KoumokuID  As Integer
  GroupID    As Integer
  CH         As Integer
  ini        As Double
  fact       As Double
  Name       As String
End Type
Public g_cTable(TDS_CH) As ct

Public Sub ReadcTable()

    Dim i As Integer, j As Integer
    Dim sa As Variant, sb As String
    Dim f As Integer
    f = FreeFile
    Open g_kankyoPath & "ctable.csv" For Input Shared As #f
    j = 0
    Do While Not (EOF(f))
        Line Input #f, sb
        Select Case Left$(sb, 1)
        Case ";", ":", "'", "#"
        Case Else
            sa = Split(sb, ",")
            j = j + 1
'            For i = 0 To UCase(sb)
                g_cTable(j).id = sa(0)
                g_cTable(j).CH = sa(1)
                g_cTable(j).ini = sa(2)
                g_cTable(j).fact = sa(3)
                
                g_cTable(j).Field = sa(5)
                g_cTable(j).KoumokuID = sa(6)
                g_cTable(j).GroupID = sa(7)
                g_cTable(j).Name = sa(9)
'            Next i
        End Select
    Loop
    Close #f
    g_cTable(0).id = j
End Sub



