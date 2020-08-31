Attribute VB_Name = "MainModule"
Option Explicit

Public Declare Function SetCurrentDirectory Lib "kernel32" Alias _
                           "SetCurrentDirectoryA" (ByVal CurrentDir As String) As Long


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public rSettei As Boolean
Public TdsDataPath(2) As String    'TDS     のデータが格納されているPath
'Public LZDataPath As String     'レーザーのデータが格納されているPath

Public BkeiPath As String
Public NoRAS As Integer

Public mINIfile As String

'メール用
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
'---------------------------------------------------------------
'---------------------------------------------------------------
Public Const PI As Double = 3.1415926535898

Public Const CTABLE_DAT As String = "ctable.dat" '環境設定ﾃﾞｰﾀ

'フォルダ名
Public TDSFTPpath(2) As String
Public PNGpath As String
Public ServerIP As String

Public CurrentDir As String    'keisoku.exeがあるフォルダ
Public Type Kei1
    Tabl_path     As String    '共通ファイルフォルダ名(danmen.dat ctable.dat kou.dat)
    keihou_path   As String    'keihou.1 keihou.2を出力するフォルダ
    Data_path     As String    'データフォルダ名
    back_path     As String
''    Theni_path    As String
    KeihouSet    As Integer
    KeihouSW     As Integer
    FH            As Integer
    LH            As Integer
''    R_ptr(1)      As Integer   '次回書き込みレコードの設定
''    R_rec(1)      As Integer   'ring.dat総レコード番号
    Waittime      As Integer   '３回計測するときの待ち時間を設定
    KeisokuCO     As Integer
''    PowerTime     As Integer
''    batteryCH     As Integer
''    batteryName   As String * 10
''    batteryUnit   As String * 10
''    batteryKei    As Single
''    batteryDec    As Integer
    CrtTpye As Integer
    crtTI As String
End Type
Public KEISOKU As Kei1

'計測インターバル設定
Public Keisoku_Time As Date '次回計測
Public KE_intv As Date      '計測インターバル
Public Z_Keisoku_Time As Date '前回計測

'2000/5/14
Public Kiroku_Time As Date    '次回記録
Public KI_intv As Date        '記録インターバル
Public Z_Kiroku_Time As Date  '前回記録
'---------

'管理値設定
Public keihou_L1 As Boolean    '警報設定管理値１
Public keihou_L2 As Boolean    '警報設定管理値２
Public keihou_L3 As Boolean    '警報設定管理値３
Public keihou_L4 As Boolean    '警報設定管理値４
Public keihou_OFF As Date      '警報を停止する時間

Public Type Kanri1
    LebelPlus(6) As Double '管理値
    LebelMinus(6) As Double '管理値
    KeihouSW As Integer '警報設定  1=True 0=False
    kanriSW As Integer  '管理設定  1=True 0=False
    sw As Integer
End Type

Public KanriMAX As Integer
Public KanriName(6) As String

'フォームを閉じた時に、データの再設定・再描画をするためのキーワード
Public Type frm1
    setTABLE  As Boolean
    setKanri  As Boolean
'''    keijiSet  As Boolean
'''    sinHosei  As Boolean
    bunpuScl  As Boolean
    setDanmen As Boolean
    setKeihouT As Boolean
    setKeihou As Boolean
    setKussaku As Boolean
    setSuiHosei  As Boolean
End Type
Public frmCLOSE As frm1

'プリンタドライバがインストールされているかどうかのチェック
Public PrntDrvSW As Boolean


Public KeihouFile As String

Public keisyaCo As Integer
Public keisyaName(10)  As String
Public keisyaR(10)  As String
Public keisyaKanri(10)  As String

Public SuiiCo As Integer
Public SuiiName(1)  As String
Public Suii(1)  As String
Public SuiiKanri(1)  As String

Public KiriBCo As Integer
Public kiribariName(10)  As String
Public kiribari(10)  As String
Public kiribariKanri(10)  As String

Public LOGFILE As String

Public Sub Main()
    Dim f As Integer, L As String
    Dim i As Integer, t_ID As Integer, d_ID As Integer
    Dim keisoku_f As Boolean
    Dim kiroku_f As Boolean
    Dim stat As Integer
    Dim Thistime As String
    Dim retDIO As Integer, retGPIB As Long, RetString As String
    Dim SS(3) As String
    Dim t1 As Date
    Dim t2 As Date
    Dim sTMP As String
    Dim sa As Variant, sb As String
    
    
    If App.PrevInstance = True Then
        MsgBox "既に起動しています。", vbCritical, "起動エラー"
        End
    End If
    
    If Command$ = "" Then
        'On Error GoTo TRAP_Main
    End If
    
100
'    ChDrive App.Path
'    ChDir App.Path
    SetCurrentDirectory App.Path
    
    CurrentDir = App.Path
    If Right(CurrentDir, 1) = "\" Then Else CurrentDir = CurrentDir & "\"
    
    mINIfile = CurrentDir & "kmSoushin.ini"
    
    '起動ログ
'    f = FreeFile
'    Open CurrentDir & "PRG-event.log" For Append Lock Write As #f
'        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 起動"
'    Close #f
    
    'ファイル読み込み
'    KEISOKU.Data_path = GetIni("フォルダ名", "計測データ", mINIfile)
'    KEISOKU.Tabl_path = GetIni("フォルダ名", "環境データ", mINIfile)
    KEISOKU.keihou_path = GetIni("フォルダ名", "警報データ", mINIfile)
    g_kankyoPath = GetIni("フォルダ名", "環境データ", mINIfile)
        
    LOGFILE = GetIni("system", "logfile", mINIfile)

GoTo 111

     keisyaCo = GetIni("keisyaFile", "CO", mINIfile)
     For i = 1 To keisyaCo
         sb = GetIni("keisyaFile", "kanri" & i, mINIfile)
         sa = Split(sb, ",")
         keisyaName(i) = sa(0)
         keisyaR(i) = sa(1)
         keisyaKanri(i) = sa(2)
     Next i

     SuiiCo = GetIni("SuiiFile", "CO", mINIfile)
     For i = 1 To SuiiCo
         sb = GetIni("SuiiFile", "kanri" & i, mINIfile)
         sa = Split(sb, ",")
         SuiiName(i) = sa(0)
         Suii(i) = sa(1)
         SuiiKanri(i) = sa(2)
     Next i

     KiriBCo = GetIni("KiriBariFile", "CO", mINIfile)
     For i = 1 To KiriBCo
         sb = GetIni("KiriBariFile", "kanri" & i, mINIfile)
         sa = Split(sb, ",")
         kiribariName(i) = sa(0)
         kiribari(i) = sa(1)
         kiribariKanri(i) = sa(2)
     Next i

111
'    TdsDataPath(1) = GetIni("フォルダ名", "データ1", mINIfile)
'    TdsDataPath(2) = GetIni("フォルダ名", "データ2", mINIfile)
    'TDSFTPpath = GetIni("フォルダ名", "TDSFTPpath", mINIfile)
'    TDSFTPpath(1) = GetIni("フォルダ名", "TDSFTPpath1", mINIfile)
'    TDSFTPpath(2) = GetIni("フォルダ名", "TDSFTPpath2", mINIfile)
    
'    ServerIP = GetIni("フォルダ名", "ServerIP", mINIfile)
    
'    PNGpath = GetIni("フォルダ名", "PNG", mINIfile)
'    BkeiPath = GetIni("フォルダ名", "Bkei", mINIfile)

    KeihouFile = GetIni("メール送信", "FileName", mINIfile)
    
    With MailTabl
        sTMP = GetIni("メール送信", "サーバー名", mINIfile): If sTMP <> "" Then .ServerName = sTMP Else MsgBox "メール送信-サーバー名が未設定", vbCritical, "設定エラー": End
        .Clientname = GetIni("メール送信", "アカウント名", mINIfile)
        .ClientMailAddress = GetIni("メール送信", "メールアドレス", mINIfile)
        .ClientRealName = GetIni("メール送信", "名前", mINIfile)
        .mailPassword = GetIni("メール送信", "パスワード", mINIfile)
        .savefolder = GetIni("メール送信", "保存場所", mINIfile)
'        .JyusinSW = CInt(GetIni("メール送信", "受信スイッチ", mINIfile))
'        .SendCO = CInt(GetIni("メール送信", "送信数", mINIfile))
'        For i = 1 To .SendCO
'            .SendName(i) = GetIni("メール送信", "送信先" & CStr(i), mINIfile)
'        Next i
    End With
    
'    Call setteiKeisya
    
'''    '2000/5/14
'''    Z_Kiroku_Time = CDate(GetIni("記録時間", "前回記録時間", mINIfile))
'''    Kiroku_Time = CDate(GetIni("記録時間", "次回記録時間", mINIfile))
'''    KI_intv = CDate(GetIni("記録時間", "記録インターバル", mINIfile))
'''    For i = 1 To 1 'DanSet(1, 0).dan
'''        KEISOKU.R_ptr(i) = CInt(GetIni("変位速度計算ファイル", "レコード", mINIfile))
'''    Next i
'''    '---------
    
    MainForm.Show
    
Exit Sub

TRAP_Main:
    f = FreeFile
    Open CurrentDir & "PRG-err.dat" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : (" & Err.Number & ") " & Err.Description
    Close #f
    
    If Command$ = "" Then
'        If KEISOKU.KeihouSW = 1 Then retDIO = MainForm.AcxDio1.Close
'        retGPIB = GpibEnd(MainForm.AcxGpib1, RetString)
        Resume 100
    End If
    End
End Sub

Public Function T_ajt(ByVal d As Date, ByVal t As Date) As Date
'次回計測日時の調整
    Dim nt As Date
    Dim t1 As Date
    Dim t2 As Date
 
    nt = d + t
    t1 = Now
    If DateDiff("s", nt, Now) > 0 Then   'If nt < Now Then
'        nt = Date + TimeSerial(Hour(Now), 0, 0)
        Do
            nt = nt + t
            t2 = Now
            If DateAdd("s", 15, t1) < t2 Then Exit Do
        Loop Until DateDiff("s", Now, nt) > 0   'nt > Now
    End If
    T_ajt = nt
End Function

Public Sub ShellSort(Aray() As Date, Num As Integer)
   Dim Span As Integer
   Dim i As Integer
   Dim j As Integer
   Dim TMP As String
   
   Span = Num \ 2
   Do While Span > 0
      For i = Span To Num - 1
         j% = i% - Span + 1
         For j = (i - Span + 1) To 1 Step -Span
            If Aray(j) >= Aray(j + Span) Then Exit For
            ' 順番の異なる配列要素を入れ替えます.
            TMP = Aray(j)
            Aray(j) = Aray(j + Span)
            Aray(j + Span) = TMP
         Next j
      Next i
      Span = Span \ 2
   Loop
End Sub

Public Sub IntvWrite()
    Call WriteIni("計測時間", "前回計測時間", Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss"), mINIfile)
    Call WriteIni("計測時間", "次回計測時間", Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss"), mINIfile)
    Call WriteIni("計測時間", "計測インターバル", Format$(KE_intv, "hh:nn:ss"), mINIfile)
'''    Call WriteIni("記録時間", "前回記録時間", Format$(Z_Kiroku_Time, "yyyy/mm/dd hh:nn:ss"), mINIfile)
'''    Call WriteIni("記録時間", "次回記録時間", Format$(Kiroku_Time, "yyyy/mm/dd hh:nn:ss"), mINIfile)
'''    Call WriteIni("記録時間", "記録インターバル", Format$(KI_intv, "hh:nn:ss"), mINIfile)

'    Open "intv.dat" For Output Access Write As #1
'    Print #1, Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'    Print #1, Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'    Print #1, Format$(KE_intv, "hh:nn:ss")
'    Close (1)
End Sub

Public Function FileCheck(FILENAME As String, FileTitle As String) As Integer
    Dim i As Integer

    On Error Resume Next

    i = 0
    If Dir$(FILENAME) = "" Then Else i = 1
    If i = 0 Then
        MsgBox FileTitle & "ファイル(" & FILENAME & ")が見つかりません。確認してください。", vbCritical, "エラーメッセージ"
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
    Dim SS As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: SS = ""
    For i = 1 To lngCheckSize
        'StrConvでUnicodeからANSIへと変換
        lngANSIStr = LenB(StrConv(Mid$(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            SS = SS + Mid$(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = SS
End Function

'Private Sub tuushin()
'    Dim t1 As Date, t2 As Date
'    Dim ret As Integer
'        With MainForm
'                        t1 = Now
'                        Do While DoEvents
'                            '回線接続まで待機
'                            If ConnectCK = 1 Then Exit Do
'                            If .RasClient1.Active = True Then ConnectCK = 1: Exit Do
'                            t2 = Now
'                            '600sec待ってもイベントが起きなければ作業中止
'                            If DateDiff("s", DateAdd("s", 600, t1), t2) > 0 Then
'                                .RasClient1.HangUp  '電話を切る
'                                Call Sleep(2000)
'                                ConnectCK = 0
'                                Exit Do
'                           '     fDial = False
'                            End If
'                        Loop
'                        If ConnectCK = 1 Then
'                        End If
'    End With
'End Sub

Public Sub AllFileDelete()
'    Dim hFso As FileSystemObject
'    Set hFso = New FileSystemObject
'
'    ' Folder オブジェクトを取得する
'    Dim hFolder As Folder
'    Set hFolder = hFso.GetFolder(TdsDataPath)
'
'    ' 不要になった時点で参照を解放する
'    Set hFso = Nothing
'
'    Dim stPrompt As String
'    Dim hFile    'As File
'
'    Dim i%
'    Dim sFile$
'    ' すべてのファイルを列挙する
'    For Each hFile In hFolder.Files
'        sFile = hFile.Path
'        Call DelFile(sFile)
'    Next hFile
'
'    ' 不要になった時点で参照を解放する
'    Set hFolder = Nothing
''    Set hFile = Nothing


End Sub

Public Sub DelFile(ByVal na As String)
    Dim hFso As FileSystemObject
    Set hFso = New FileSystemObject

    On Local Error Resume Next
    ' 読み取り専用ファイルの場合でも削除する
    Call hFso.DeleteFile(na, True)

    ' 不要になった時点で参照を解放する
    Set hFso = Nothing
End Sub

Public Sub ErrLog(ByVal d As Date, ByVal t As String, ByVal s As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\err_log.txt" For Append As #f
    Print #f, Format$(d, ""); vbTab;
    Print #f, t; vbTab;
    Print #f, s
    Close #f
End Sub

Private Sub StatusMessage(ByVal sa As String)
    MainForm.StatusBar1.Panels(1).Text = sa
End Sub

'###標準モジュールに
Public Sub SendPNG(ret As Integer)
'', fdir As String, fPath() As String, FTPpath$)
'
'Dim SendPath, rSettei
''リモートへファイルを送信します｡複数ファイルの送信ができます｡
''
''rc = ftp.PutFile(local,remote[,type])
''  local [in]  : 送信するファイル名をフルパスで指定。
''                複数ファイルの指定は、 "a*.txt" 、"*"、"*.html" などのように "*" を使う。
''                例： c:\html\a.html --- htmlディレクトリのa.html
''                     c:\html\*.html --- htmlディレクトリの .html ファイルすべて
''                     c:\html\*      --- htmlディレクトリのすべてのファイル
'' remote [in]  : リモートのディレクトリ名。"" は、カレントディレクトリ。
'' Enum in: 送信するデータ形式を次のように指定｡
''  0 : ASCII（省略値)。txt/html などのテキストファイルの場合。
''  1 : バイナリ。jpg/gif/exe/lzh/tar.gz などのバイナリファイルの場合。
''  2 : ASCII + 追加(Append)モード。
''  3 : バイナリ + 追加(Append)モード。
''
''  rc [out]: 結果コードが数字で返されます｡
''  1 以上:   正常終了｡送信したファイル数｡
''  0     :   該当するファイルなし｡
''  1,0以外 : エラー。GetReplyメソッドでFTP応答テキストで詳細を調べてください。
''例:
''rc = ftp.PutFile("c:\html\index.html", "html")  ' テキストファイルの送信
''rc = ftp.PutFile("c:\html\*.html", "html")      ' テキストファイルの送信
''rc = ftp.PutFile("c:\html\*.html", "html", 2)     ' テキストファイルのAppendモード送信
''rc = ftp.PutFile("c:\html\images\*", "html/images", 1) ' バイナリファイルの送信
'
'    Dim i As Integer
'    Dim tFile As String
'
'    Dim sYY As String
'    Dim sMM As String
'    Dim sDD As String
'    Dim fpSW As Boolean
'    ret = 0
'    On Local Error GoTo SendPNGerr
'
'    Dim ftpErr  As String
'    Dim rc As Long
'    Dim vv As Variant, vv2 As Variant
'''    Dim ftp As Object
'''    Set ftp = CreateObject("basp21.FTP")
'    Dim ftp As BASP21Lib.ftp
'    Set ftp = New BASP21Lib.ftp
'
'    ftp.OpenLog App.Path & "\FTP-log-png.txt"
'    'rc = ftp.Connect("172.16.60.99", "anonymous", "")  'TEST
'    rc = ftp.Connect("60.43.239.36", "chikah", "zbeba+nn")  '本物
'    If rc = 0 Then
'        'データのアップロード
'        rc = ftp.PutFile(PNGpath & "*", "/pic", 1) 'ファイル送信
'
'        ftp.Close
'    Else
'        ftpErr = ftp.GetReply()
'    End If
'    rc = ftp.CloseLog()
'
'    Set ftp = Nothing
'    ret = -1
'Exit Sub
'SendPNGerr:
'    Set ftp = Nothing
'    ret = 0
End Sub


Public Sub WriteLog(st As String)
'st 説明文
    Dim f As Integer
    
    On Error GoTo WriteLog9999
    
    f = FreeFile
    Open App.Path & "\" & LOGFILE For Append As #f
    Print #f, Format(Now, "YYYY/MM/DD hh:mm:ss"); " : ";
    Print #f, st
    Close #f

WriteLog9999:
    On Error GoTo 0
End Sub

