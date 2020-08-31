Attribute VB_Name = "FileRead"
Option Explicit

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
'   断面ファイル
'**********************************************************************************************
Public Sub ReadDan()
    Dim f As Integer
    Dim L As String
    Dim i As Integer
    Dim k_ID As Integer
    Dim d_ID As Integer
    
    i = FileCheck(KEISOKU.Tabl_path & "danmen.dat", "環境データ")
    If i = 0 Then Unload MainForm: End
    
    Erase DanSet
    
    f = FreeFile
    Open KEISOKU.Tabl_path & "danmen.dat" For Input Access Read Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L
        If Left$(L, 1) <> ":" Then
            k_ID = CInt(Mid$(L, 1, 4))
            d_ID = CInt(Mid$(L, 5, 4))
            DanSet(k_ID, d_ID).Ti = Trim$(SEEKmoji(L, 9, 12))
            DanSet(k_ID, d_ID).Plus = Trim$(SEEKmoji(L, 21, 10))
            DanSet(k_ID, d_ID).Minus = Trim$(SEEKmoji(L, 31, 10))
            DanSet(k_ID, d_ID).kou = k_ID
            DanSet(k_ID, d_ID).dan = d_ID
            
            If DanSet(k_ID, 0).dan < d_ID Then DanSet(k_ID, 0).dan = d_ID
        End If
    Loop
    Close (f)
End Sub

'**********************************************************************************************
'   初期値・係数・ﾁｬﾝﾈﾙ・作図測点などのTABLE.set
'**********************************************************************************************
Public Sub ReadTabl()
    Dim f As Integer
    Dim i As Integer
    Dim L As String
    Dim d_ID As Integer
    Dim k_ID As Integer, t_ID As Integer
'''    Dim FLDno As Integer
    
    Erase Tbl
    
    i = FileCheck(KEISOKU.Tabl_path & CTABLE_DAT, "環境データ")
    If i = 0 Then Unload MainForm: End
    
    TDSTbl(0).ch = 0
    For i = 1 To TDS_CH: TDSTbl(i).ch = 999: Next i
    
    d_ID = 1
    i = 0
    f = FreeFile
    Open KEISOKU.Tabl_path & CTABLE_DAT For Input Access Read Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L
        Select Case Left$(L, 1)
        Case ":", ";", "#"
        Case Else
            i = i + 1
            k_ID = CInt(Mid$(L, 5, 4))
            t_ID = CInt(Mid$(L, 9, 4))
            
            Tbl(k_ID, d_ID, t_ID).kou = k_ID
            Tbl(k_ID, d_ID, t_ID).ten = t_ID
            Tbl(k_ID, d_ID, t_ID).FLD = CInt(Mid$(L, 1, 4))
            Tbl(k_ID, d_ID, t_ID).ch = CInt(Mid$(L, 13, 4))
            Tbl(k_ID, d_ID, t_ID).Syo = CDbl(Mid$(L, 17, 10))
            Tbl(k_ID, d_ID, t_ID).Kei = CDbl(Mid$(L, 27, 10))
            Tbl(k_ID, d_ID, t_ID).HAN = Trim$(SEEKmoji(L, 37, 8))
            
            Tbl(k_ID, d_ID, 0).ten = Tbl(k_ID, d_ID, 0).ten + 1
            
            TDSTbl(i).ch = CInt(Mid$(L, 13, 4))
            TDSTbl(i).kou = k_ID
            TDSTbl(i).FLD = Tbl(k_ID, d_ID, t_ID).FLD
        End Select
    Loop
    Close #f
    TDSTbl(0).ch = i
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
            
            kou(k_ID, s_ID).no = k_ID
            kou(k_ID, s_ID).KIND = s_ID
            
            If kou(0, 1).no < k_ID Then kou(0, 1).no = k_ID
            kou(k_ID, 0).no = kou(k_ID, 0).no + 1
        End Select
    Loop
    Close #f
End Sub

