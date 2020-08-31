Attribute VB_Name = "SyokiModule"
Option Explicit

Public MDY
Public Const MAX_CH  As Integer = 27   '      〃     のデータ数
Public CurrentDir As String   'keisoku.exeがあるフォルダ
Public Tabl_path As String

Public SyokiSW As Boolean
Public PoSW As Boolean

'環境設定
Public Const Maxkou As Integer = 1     '計測項目数
Public Const Maxdan As Integer = 1     '計測断面数
Public Type table1
    dan   As Integer     '断面番号
    kou   As Integer     '項目番号
    ten   As Integer     '測点番号
    FLD   As Integer     'データファイルフィールド位置
    ch    As Integer     '計測チャンネル
    Syo   As Double      '初期値
    Kei   As Double      '係数
    deep  As Single      '深度
    HAN   As String * 8  '凡例
    keiji As Integer     '経時変化図作図測点 1.する 0.しない
End Type
Public Tbl(Maxkou, Maxdan, 10) As table1

Public Type RsInit1
    DeviceNo As Integer
    SpdNO As Integer
    PrtNO As Integer
    sizeNO As Integer
    stopNo As Integer
    Stime As Long
    Rtime As Long
End Type
Public RsInit As RsInit1


Public Type Init1
    Serch As String
    Wait As String
    PoFILE0 As String
    PoFILE1 As String
    Tensu As Integer
    HeikinKaisuu As Integer
    x0 As Double
    y0 As Double
    z0 As Double
    MH As Double
    x1 As Double
    y1 As Double
    z1 As Double
    AZIMUTH As Single
'    CO As Integer
'    AvgCO As Integer
'    Kx As Double
'    Ky As Double
'    Kz As Double
'    Kmh As Double
'    Bx As Double
'    By As Double
'    Bz As Double
'    HOKO As Single
End Type
Public InitDT As Init1

Public Type Po1
    H(2, 16) As Double
    V(2, 16) As Double
    s(2, 16) As Double
'    Hdt As Double
'    Vdt As Double
'    Sdt As Double
End Type
Public PoDT As Po1
Public RsctlFrm As Object
'フォームを閉じた時に、データの再設定・再描画をするためのキーワード
Public Type frm1
    setTABLE As Boolean
    setKanri As Boolean
    keijiSet As Boolean
    bunpuScl As Boolean
    sinHosei As Boolean
    StartFrm As Boolean
    MSGfrm As Boolean
End Type
Public frmCLOSE As frm1

Public DH(17) As Double
Public DV(17) As Double
Public XD(20) As Double '測定座標
Public YD(20) As Double
Public ZD(20) As Double
Public xo(16) As Double '測定座標（前値）
Public yo(16) As Double
Public zo(16) As Double

Public Const RAD As Double = 3.14159265358979 / 180#



Sub Main()
    Dim rc As Integer
    Dim i As Integer
    
    If App.PrevInstance = True Then
        MsgBox "既に起動しています。", vbCritical, "起動エラー"
        End
    End If
    
'    ChDrive App.Path
'    ChDir App.Path
    SetCurrentDirectory App.Path
    
    CurrentDir = App.Path
    If Right(CurrentDir, 1) = "\" Then Else CurrentDir = CurrentDir & "\"
    
    Set RsctlFrm = frmPOset
    
    Tabl_path = GetIni("フォルダ名", "環境データ", CurrentDir & "計測設定.ini")
    Call ReadTabl
    
    '計測条件
    InitDT.Serch = GetIni("計測条件", "サーチ設定", CurrentDir & "計測設定.ini")
    InitDT.Wait = GetIni("計測条件", "ウェイト時間設定", CurrentDir & "計測設定.ini")
    InitDT.PoFILE0 = GetIni("計測条件", "測点位置ファイル初期値", CurrentDir & "計測設定.ini")
    InitDT.PoFILE1 = GetIni("計測条件", "測点位置ファイル前回値", CurrentDir & "計測設定.ini")
    
    InitDT.Tensu = CInt(GetIni("計測条件", "測点数", CurrentDir & "計測設定.ini"))
    InitDT.x0 = CDbl(GetIni("計測条件", "器械点X", CurrentDir & "計測設定.ini"))
    InitDT.y0 = CDbl(GetIni("計測条件", "器械点Y", CurrentDir & "計測設定.ini"))
    InitDT.z0 = CDbl(GetIni("計測条件", "器械点Z", CurrentDir & "計測設定.ini"))
    InitDT.MH = CDbl(GetIni("計測条件", "器械点MH", CurrentDir & "計測設定.ini"))
    InitDT.x1 = CDbl(GetIni("計測条件", "後視点X", CurrentDir & "計測設定.ini"))
    InitDT.y1 = CDbl(GetIni("計測条件", "後視点Y", CurrentDir & "計測設定.ini"))
    InitDT.z1 = CDbl(GetIni("計測条件", "後視点Z", CurrentDir & "計測設定.ini"))
    InitDT.AZIMUTH = CDbl(GetIni("計測条件", "方向角", CurrentDir & "計測設定.ini"))
    
'    i = FileCheck(InitDT.PoFILE, "計測位置ファイル")
'    If i = 0 Then End
   
    '通信設定
    RsInit.DeviceNo = CInt(GetIni("通信設定", "通信ポート", CurrentDir & "計測設定.ini"))
    RsInit.SpdNO = CInt(GetIni("通信設定", "通信速度", CurrentDir & "計測設定.ini"))
    RsInit.PrtNO = CInt(GetIni("通信設定", "パリティ", CurrentDir & "計測設定.ini"))
    RsInit.sizeNO = CInt(GetIni("通信設定", "データ長", CurrentDir & "計測設定.ini"))
    RsInit.stopNo = CInt(GetIni("通信設定", "ストップ", CurrentDir & "計測設定.ini"))
    RsInit.Rtime = CLng(GetIni("通信設定", "受信タイムアウト", CurrentDir & "計測設定.ini"))
    RsInit.Stime = CLng(GetIni("通信設定", "送信タイムアウト", CurrentDir & "計測設定.ini"))
    
    frmSyokiset.Show

End Sub

Public Function FileCheck(FileName As String, FileTitle As String) As Integer
    Dim i As Integer

    On Error Resume Next

    i = 0
    If Dir$(FileName) = "" Then Else i = 1
    If i = 0 Then
        MsgBox FileTitle & "ファイル(" & FileName & ")が見つかりません。確認してください。", vbCritical, "エラーメッセージ"
    End If
    
    FileCheck = i
    
    On Error GoTo 0

End Function

Public Function SEEKmoji(strCheckString As String, mojiST As Integer, mojiMAX As Integer) As String

    'Forカウンタ
    Dim i As Long
    '調査対象文字列の長さを格納
    Dim lngCheckSize As Long
    'ANSIへの変換後の文字を格納
    Dim lngANSIStr As Long
    
    Dim co As Integer '文字数
    Dim ss As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: ss = ""
    For i = 1 To lngCheckSize
        'StrConvでUnicodeからANSIへと変換
        lngANSIStr = LenB(StrConv(Mid$(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            ss = ss + Mid$(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = ss
End Function


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
    
    i = FileCheck(Tabl_path & "CTABLE.DAT", "環境データ")
    If i = 0 Then End
    
    d_ID = 1
    i = 0
    f = FreeFile
    Open Tabl_path & "CTABLE.DAT" For Input Shared As #f
    Do While Not (EOF(f))
        Line Input #f, L
        If Left$(L, 1) <> ";" Then
            i = i + 1
            k_ID = CInt(Mid$(L, 5, 4))
            t_ID = CInt(Mid$(L, 9, 4))
            
            Tbl(k_ID, d_ID, t_ID).kou = k_ID
            Tbl(k_ID, d_ID, t_ID).ten = t_ID
            Tbl(k_ID, d_ID, t_ID).FLD = CInt(Mid$(L, 1, 4))
            Tbl(k_ID, d_ID, t_ID).ch = CInt(Mid$(L, 13, 4))
            Tbl(k_ID, d_ID, t_ID).Syo = CSng(Mid$(L, 17, 10))
            Tbl(k_ID, d_ID, t_ID).Kei = CDbl(Mid$(L, 27, 10))
            Tbl(k_ID, d_ID, t_ID).HAN = Trim$(SEEKmoji(L, 37, 8))
            
            Tbl(k_ID, d_ID, 0).ten = Tbl(k_ID, d_ID, 0).ten + 1
            
        End If
    Loop
    Close #f
End Sub


