Attribute VB_Name = "MainModule"
Option Explicit

Public Declare Function SetCurrentDirectory Lib "kernel32" Alias _
                           "SetCurrentDirectoryA" (ByVal CurrentDir As String) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub DisableProcessWindowsGhosting Lib "user32" ()

Public rSettei As Boolean
Public TdsDataPath(10) As String    'TDS     のデータが格納されているPath

Public BkeiPath As String
Public NoRAS As Integer

'---------------------------------------------------------------
'---------------------------------------------------------------
Public Const PI As Double = 3.1415926535898

Public Const CTABLE_DAT As String = "ctable.dat" '環境設定ﾃﾞｰﾀ

'フォルダ名
Public TDSFTPpath(10) As String
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

Public Atesaki As String, sUser As String
Public fco As Integer
    
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
    
    
    If App.PrevInstance = True Then
''        MsgBox "既に起動しています。", vbCritical, "起動エラー"
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
    
    mINIfile = CurrentDir & "fSoushin.ini"
    
    '起動ログ
    f = FreeFile
    Open CurrentDir & App.EXEName & "-event.log" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 起動"
    Close #f
    
    'ファイル読み込み
'    KEISOKU.Data_path = GetIni("フォルダ名", "計測データ", mINIfile)
'    KEISOKU.Tabl_path = GetIni("フォルダ名", "環境データ", mINIfile)
    KEISOKU.keihou_path = GetIni("フォルダ名", "警報データ", mINIfile)
    fco = GetIni("フォルダ名", "co", mINIfile)
    For i = 1 To fco
        TdsDataPath(i) = GetIni("フォルダ名", "データ" & i, mINIfile)
        TDSFTPpath(i) = GetIni("フォルダ名", "FTPpath" & i, mINIfile)
    Next i
    
    'ServerIP = GetIni("フォルダ名", "ServerIP", mINIfile)
    
    'PNGpath = GetIni("フォルダ名", "PNG", mINIfile)
    'BkeiPath = GetIni("フォルダ名", "Bkei", mINIfile)

'####################################
    Atesaki = "153.150.115.38"
    sUser = "onuma"
'####################################
    DisableProcessWindowsGhosting
    
    MainForm.Show
    
Exit Sub

TRAP_Main:
    f = FreeFile
    Open CurrentDir & App.EXEName & "-err.txt" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : (" & Err.Number & ") " & Err.Description
    Close #f
    
    If Command$ = "" Then
'        If KEISOKU.KeihouSW = 1 Then retDIO = MainForm.AcxDio1.Close
'        retGPIB = GpibEnd(MainForm.AcxGpib1, RetString)
        Resume 100
    End If
    End
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
    Open App.Path & App.EXEName & "-err_log.txt" For Append As #f
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
', fdir As String, fPath() As String, FTPpath$)

Dim SendPath, rSettei
'リモートへファイルを送信します｡複数ファイルの送信ができます｡
'
'rc = ftp.PutFile(local,remote[,type])
'  local [in]  : 送信するファイル名をフルパスで指定。
'                複数ファイルの指定は、 "a*.txt" 、"*"、"*.html" などのように "*" を使う。
'                例： c:\html\a.html --- htmlディレクトリのa.html
'                     c:\html\*.html --- htmlディレクトリの .html ファイルすべて
'                     c:\html\*      --- htmlディレクトリのすべてのファイル
' remote [in]  : リモートのディレクトリ名。"" は、カレントディレクトリ。
' Enum in: 送信するデータ形式を次のように指定｡
'  0 : ASCII（省略値)。txt/html などのテキストファイルの場合。
'  1 : バイナリ。jpg/gif/exe/lzh/tar.gz などのバイナリファイルの場合。
'  2 : ASCII + 追加(Append)モード。
'  3 : バイナリ + 追加(Append)モード。
'
'  rc [out]: 結果コードが数字で返されます｡
'  1 以上:   正常終了｡送信したファイル数｡
'  0     :   該当するファイルなし｡
'  1,0以外 : エラー。GetReplyメソッドでFTP応答テキストで詳細を調べてください。
'例:
'rc = ftp.PutFile("c:\html\index.html", "html")  ' テキストファイルの送信
'rc = ftp.PutFile("c:\html\*.html", "html")      ' テキストファイルの送信
'rc = ftp.PutFile("c:\html\*.html", "html", 2)     ' テキストファイルのAppendモード送信
'rc = ftp.PutFile("c:\html\images\*", "html/images", 1) ' バイナリファイルの送信
    
    Dim i As Integer
    Dim tFile As String
    
    Dim sYY As String
    Dim sMM As String
    Dim sDD As String
    Dim fpSW As Boolean
    ret = 0
    On Local Error GoTo SendPNGerr
    
    Dim ftpErr  As String
    Dim rc As Long
    Dim vv As Variant, vv2 As Variant
''    Dim ftp As Object
''    Set ftp = CreateObject("basp21.FTP")
    Dim ftp As BASP21Lib.ftp
    Set ftp = New BASP21Lib.ftp
    
    ftp.OpenLog App.Path & "\FTP-log-png.txt"
    'rc = ftp.Connect("172.16.60.99", "anonymous", "")  'TEST
    rc = ftp.Connect("60.43.239.36", "chikah", "zbeba+nn")  '本物
    If rc = 0 Then
        'データのアップロード
        rc = ftp.PutFile(PNGpath & "*", "/pic", 1) 'ファイル送信
        
        ftp.Close
    Else
        ftpErr = ftp.GetReply()
    End If
    rc = ftp.CloseLog()
    
    Set ftp = Nothing
    ret = -1
Exit Sub
SendPNGerr:
    Set ftp = Nothing
    ret = 0
End Sub

