Attribute VB_Name = "FileRead"
Option Explicit

'環境設定
Public Type table1
    id    As Integer     '断面番号
    dan   As Integer     '断面番号
    kou   As Integer     '項目番号
    danme As Integer
    muki  As Integer     '向き
    ten   As Integer     '測点番号
    po    As Integer
        
    FLD   As Integer     'データファイルフィールド位置
    TDSid As Integer     'ＴＤＳ No.
    CH    As Integer     '計測チャンネル
    TDSsw As Integer     '0=ＴＤＳからきたデータを無視して欠測とする。
    Syo   As Double      '初期値
    Kei   As Double      '係数
    Name       As String
    hanare     As Double      '基準点からの距離 mm
    SyoOndo    As Double      '初期値のときの温度
    OndoKei    As Double      '温度係数
    
    deep       As Double      '深度
    leng       As Double
    grNo  As Integer     '
    Dkei       As Double
    Dseki      As Double
    DHAN       As String   '凡例
    HAN        As String   '凡例
    keikiName  As String   '凡例
'''    Tch   As Integer
'''    ThoseiA As double
'''    ThoseiB As double
'''    ThoseiT As double
'''    keiji As Integer     '経時変化図作図測点 1.する 0.しない
End Type
Public tbl(200) As table1

Public Sub ReadKanri()
    Dim f As Integer
    Dim L As String
    Dim i As Integer
    Dim d_ID As Integer
    Dim l_ID As Integer
    
    i = FileCheck(KEISOKU.Tabl_path & "kanri.dat", "管理値データ")
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
'   初期値・係数・ﾁｬﾝﾈﾙ・作図測点などのTABLE.set
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
    
'    i = FileCheck(KEISOKU.Tabl_path & CTABLE_DAT, "環境データ")
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
'   項目ファイル
'**********************************************************************************************
Public Sub ReadKou()
    Dim f As Integer
    Dim i As Integer
    Dim L As String
    Dim k_ID As Integer, s_ID As Integer
    
    i = FileCheck(KEISOKU.Tabl_path & "kou.dat", "環境データ")
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
