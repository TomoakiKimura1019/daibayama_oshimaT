Attribute VB_Name = "Module1"
Option Explicit

Public g_kankyoPath As String
Public g_keisyaBTM(5) As Double
Public g_keisyaConf(5) As String
Public g_keisyaKanriF(5) As String
Public g_keisyaDepth(20) As Double

Public Sub setteiKeisya()
    Dim j As Integer
    Dim f As Integer
    Dim bf As String
    Dim sa As Variant
    Dim cp As Integer
    Dim pa1 As String
    pa1 = g_kankyoPath
    
    f = FreeFile
    Open pa1 & "�X�Όv.dat" For Input Access Read Lock Write As #f
    Input #f, cp
    Do While Not (EOF(f))
        Line Input #f, bf
        Select Case Left$(bf, 1)
        Case ";"
        Case Else
            sa = Split(bf, ",")
            j = sa(0)                         ' �E�ԍ�
            g_keisyaBTM(j) = Trim(sa(1))      ' �Ő[�ʒu(m)
            g_keisyaConf(j) = Trim(sa(2))     ' �X�̍E�̒�`�t�@�C����
            g_keisyaKanriF(j) = Trim(sa(3))     ' �X�̍E�̊Ǘ��l�t�@�C����
        End Select
    Loop
    Close #f
    
End Sub

