Attribute VB_Name = "FileRead"
Option Explicit

'���ݒ�
Public Type table1
    id    As Integer     '�f�ʔԍ�
    dan   As Integer     '�f�ʔԍ�
    kou   As Integer     '���ڔԍ�
    danme As Integer
    muki  As Integer     '����
    ten   As Integer     '���_�ԍ�
    po    As Integer
        
    FLD   As Integer     '�f�[�^�t�@�C���t�B�[���h�ʒu
    TDSid As Integer     '�s�c�r No.
    CH    As Integer     '�v���`�����l��
    TDSsw As Integer     '0=�s�c�r���炫���f�[�^�𖳎����Č����Ƃ���B
    Syo   As Double      '�����l
    Kei   As Double      '�W��
    Name       As String
    hanare     As Double      '��_����̋��� mm
    SyoOndo    As Double      '�����l�̂Ƃ��̉��x
    OndoKei    As Double      '���x�W��
    
    deep       As Double      '�[�x
    leng       As Double
    grNo  As Integer     '
    Dkei       As Double
    Dseki      As Double
    DHAN       As String   '�}��
    HAN        As String   '�}��
    keikiName  As String   '�}��
'''    Tch   As Integer
'''    ThoseiA As double
'''    ThoseiB As double
'''    ThoseiT As double
'''    keiji As Integer     '�o���ω��}��}���_ 1.���� 0.���Ȃ�
End Type
Public tbl(200) As table1

Public Sub ReadKanri()
    Dim f As Integer
    Dim L As String
    Dim i As Integer
    Dim d_ID As Integer
    Dim l_ID As Integer
    
    i = FileCheck(KEISOKU.Tabl_path & "kanri.dat", "�Ǘ��l�f�[�^")
    If i = 0 Then Exit Sub
    
    d_ID = 1
    f = FreeFile
    Open KEISOKU.Tabl_path & "kanri.dat" For Input Access Read Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L
        Select Case Left$(L, 1)
        Case ";", ":", "#"
        Case Else
            l_ID = CInt(Mid$(L, 1, 4))
            Kanri(1, d_ID).Lebel1(l_ID) = CDbl(Mid$(L, 5, 8))
            Kanri(1, d_ID).Lebel2(l_ID) = CDbl(Mid$(L, 13, 8))
            Kanri(1, d_ID).TI1(l_ID) = Trim(SEEKmoji(L, 21, 8))
            Kanri(1, d_ID).TI2(l_ID) = Trim(SEEKmoji(L, 29, 12))
        End Select
    Loop
    Close #f
    
End Sub

'**********************************************************************************************
'   �����l�E�W���E����فE��}���_�Ȃǂ�TABLE.set
'**********************************************************************************************
Public Sub ReadTabl() '(seigen() As Double, tCH() As Integer)
    Dim f As Integer
    Dim i As Integer
    Dim L As String
    Dim d_ID As Integer, dd_id As Integer, m_ID As Integer
    Dim k_ID As Integer, t_id As Integer, p_ID As Integer
    Dim FLDno As Integer
    
    Dim id As Integer
    
    Dim sa As Variant
    Dim sFileName As String
    sFileName = "cTable.csv"
    Erase tbl ', tblNO, MaxTen, MaxDanme, MaxTenK, tblNOK, MaxTenS
    
'    i = FileCheck(KEISOKU.Tabl_path & CTABLE_DAT, "���f�[�^")
'    If i = 0 Then Unload frmView1: End
    
    tbl(0).CH = 0
    For i = 0 To 100
        tbl(i).CH = 999
        tbl(i).TDSsw = 1
    Next i
    
    f = FreeFile
    Open KEISOKU.Tabl_path & sFileName For Input Shared As #f
    i = 0
    Do While Not (EOF(f))
        Line Input #f, L
        Select Case Left$(L, 1)
        Case ";", ":", "#"
        Case Else
            i = i + 1
            sa = Split(L, ",")
            id = sa(0)
            tbl(id).id = id
            tbl(id).CH = CDbl(sa(1))
            tbl(id).Syo = CDbl(sa(2))
            tbl(id).Kei = CDbl(sa(3))
            tbl(id).TDSsw = CDbl(sa(4))
            tbl(id).FLD = CDbl(sa(5))
            tbl(id).kou = CDbl(sa(6))
            tbl(id).dan = CDbl(sa(7))
            tbl(id).grNo = CDbl(sa(7))
            tbl(id).Name = (sa(9))
        End Select
    Loop
    Close #f
    
    tbl(0).id = i
'    tbl(0).FLD = i
    
End Sub

'**********************************************************************************************
'   ���ڃt�@�C��
'**********************************************************************************************
Public Sub ReadKou()
    Dim f As Integer
    Dim i As Integer
    Dim L As String
    Dim k_ID As Integer, s_ID As Integer
    
    i = FileCheck(KEISOKU.Tabl_path & "kou.dat", "���f�[�^")
    If i = 0 Then Unload MainForm: End

    Erase kou
    
    f = FreeFile
    Open KEISOKU.Tabl_path & "kou.dat" For Input Access Read Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L:
        Select Case Left$(L, 1)
        Case ":", ";", "#"
        Case Else
            k_ID = CInt(Mid$(L, 1, 4))
            s_ID = CInt(Mid$(L, 5, 4))
            
            kou(k_ID, s_ID).TI1 = Trim$(SEEKmoji(L, 9, 20))
            kou(k_ID, s_ID).TI2 = Trim$(SEEKmoji(L, 29, 20))
            kou(k_ID, s_ID).Yt = Trim$(SEEKmoji(L, 49, 10))
            kou(k_ID, s_ID).Yu = Trim$(SEEKmoji(L, 59, 10))
            kou(k_ID, s_ID).dec = CInt(SEEKmoji(L, 69, 4))
            
            kou(k_ID, s_ID).No = k_ID
            kou(k_ID, s_ID).KIND = s_ID
            
            If kou(0, 1).No < k_ID Then kou(0, 1).No = k_ID
            kou(k_ID, 0).No = kou(k_ID, 0).No + 1
        End Select
    Loop
    Close #f
End Sub
