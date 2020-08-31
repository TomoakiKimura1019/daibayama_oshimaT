Attribute VB_Name = "GpibSubRoutines"
Option Explicit

'ｷｰﾜｰﾄﾞの定義
Global Const GP_GTL As Long = &H1
Global Const GP_SDC As Long = &H4
Global Const GP_PPC As Long = &H5
Global Const GP_GET As Long = &H8
Global Const GP_TCT As Long = &H9
Global Const GP_LLO As Long = &H11
Global Const GP_DCL As Long = &H14
Global Const GP_PPU As Long = &H15
Global Const GP_SPE As Long = &H18
Global Const GP_SPD As Long = &H19
Global Const GP_MLA As Long = &H20
Global Const GP_UNL As Long = &H3F
Global Const GP_MTA As Long = &H40
Global Const GP_UNT As Long = &H5F

'Windowsで規定されている値
Global Const HELP_CONTEXT = &H1
Global Const HELP_QUIT = &H2
Global Const HELP_CONTENTS = &H3

'ACX-GPIB(W32) Helpコンテキスト
Global Const HLP_SAMPLES = 274
Global Const HLP_SAMPLES_BASIC = 275
Global Const HLP_SAMPLES_EVENT = 276
Global Const HLP_SAMPLES_MULTILINE = 277
Global Const HLP_SAMPLES_MULTIMETER = 278
Global Const HLP_SAMPLES_POLLING = 279
Global Const HLP_SAMPLES_PARARELL = 280
Global Const HLP_SAMPLES_VOLT = 281

'LoadProperty, SavePropertyで使用
Global Const ERR_FILE_NOT_FOUND = 190
Global Const ERR_FILE_COULD_NOT_OPEN = 191
Global Const ERR_FILE_WRITE = 192
Global Const ERR_FILE_READ = 193
Global Const ERR_FILE_UNKNOWN = 194
Global Const ERR_FILE_INVALID_FORMAT = 195

'Win32 APIのコールのための宣言
Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

Public Function GpibInit(Gp As Object, RetSts As String) As Long
'初期化のためのサブルーチン。オブジェクト名を引数として受け取ります。
'戻り値 : 正常終了 = 0、異常終了 = 1

Dim ret As Long

'再初期化防止のため
    Gp.Exit

    GpibInit = 0
    
'Ini、Ifc、Renの3つのメソッドで初期化のひとかたまりになります。
'ボードの初期化
    ret = Gp.Ini
    If ret <> 0 Then
        GpibInit = CheckRetCode("Ini", ret, RetSts)
        Exit Function
    End If
    
'マスタの時のみ以下の2つのメソッドを実行します
    If Gp.MasterSlave = 0 Then
    'IFC(Interface Clear)の送出
        ret = Gp.Ifc
        If ret <> 0 Then
            GpibInit = CheckRetCode("Ifc", ret, RetSts)
            Exit Function
        End If
    'リモートラインを有効にする
        ret = Gp.Ren
        If ret <> 0 Then
            GpibInit = CheckRetCode("Ren", ret, RetSts)
            Exit Function
        End If
    End If

'正常終了のときは以下の文字列を返します
    RetSts = "初期化完了"
    GpibInit = 0

End Function

Public Function GpibEnd(Gp As Object, RetSts As String) As Long
'終了のためのサブルーチン。オブジェクト名を引数として受け取ります。

Dim ret As Long

'マスタの時のみ以下の2つのメソッドを実行します
    If Gp.MasterSlave = 0 Then
    'リモートラインのリセット
        ret = Gp.Resetren
        If ret <> 0 Then
            GpibEnd = CheckRetCode("Resetren", ret, RetSts)
            Exit Function
        End If
    End If

'終了処理の実行
    ret = Gp.Exit
    If ret <> 0 Then
        GpibEnd = CheckRetCode("Exit", ret, RetSts)
        Exit Function
    End If

    RetSts = "正常終了"
    GpibEnd = 0

End Function

Public Function CheckRetCode(Buf As String, RetCode As Long, RetBuf As String) As Long
'エラーチェックサブルーチン。表示するメソッド名と戻り値を引数として受け取ります。

Dim CheckRet As Long
Dim RetSts As Long
Dim TextErr As String

    TextErr = Buf + " : 正常終了"
    CheckRet = RetCode And &HFF
    RetSts = 0
    
    If (CheckRet >= 3) Then
        RetSts = 1
        If (CheckRet = 3) Then TextErr = Buf + " : FIFO内にまだデータが残っています": GoTo CheckStatus
        If (CheckRet = 80) Then TextErr = Buf + " : I/Oアドレスエラー": GoTo CheckStatus
        If (CheckRet = 128) Then TextErr = Buf + " : データ受信予定数を超えたか(受信)またはSRQを受信していません(ポーリング)": GoTo CheckStatus
        If (CheckRet = 200) Then TextErr = Buf + " : スレッドが作成できません": GoTo CheckStatus
        If (CheckRet = 240) Then TextErr = Buf + " : Escキーが押されました": GoTo CheckStatus
        If (CheckRet = 241) Then TextErr = Buf + " : File入出力エラー": GoTo CheckStatus
        If (CheckRet = 242) Then TextErr = Buf + " : アドレス指定ミス": GoTo CheckStatus
        If (CheckRet = 243) Then TextErr = Buf + " : バッファ指定エラー": GoTo CheckStatus
        If (CheckRet = 244) Then TextErr = Buf + " : 配列サイズエラー": GoTo CheckStatus
        If (CheckRet = 245) Then TextErr = Buf + " : バッファが小さすぎます": GoTo CheckStatus
        If (CheckRet = 246) Then TextErr = Buf + " : 不正なオブジェクト名です": GoTo CheckStatus
        If (CheckRet = 247) Then TextErr = Buf + " : デバイス名の横のチェックが無効です": GoTo CheckStatus
        If (CheckRet = 248) Then TextErr = Buf + " : 不正なデータ型です": GoTo CheckStatus
        If (CheckRet = 249) Then TextErr = Buf + " : これ以上デバイスを追加できません": GoTo CheckStatus
        If (CheckRet = 250) Then TextErr = Buf + " : デバイス名が見つかりません": GoTo CheckStatus
        If (CheckRet = 251) Then TextErr = Buf + " : デリミタがデバイス間で違っています": GoTo CheckStatus
        If (CheckRet = 252) Then TextErr = Buf + " : GP-IBエラー": GoTo CheckStatus
        If (CheckRet = 253) Then TextErr = Buf + " : デリミタのみを受信しました": GoTo CheckStatus
        If (CheckRet = 254) Then TextErr = Buf + " : タイムアウトしました": GoTo CheckStatus
        If (CheckRet = 255) Then TextErr = Buf + " : パラメータエラー": GoTo CheckStatus
                TextErr = Buf + " : このサンプルではエラーコード" & CheckRet & "はサポートしていません。"
    End If

CheckStatus:
    '----- Ifc & Srq Receive Status Message ------------
    CheckRet = RetCode And &HFF00
    If (CheckRet = &H100) Then TextErr = TextErr + " -- SRQを受信しました <ステータス>": GoTo CheckEnd
    If (CheckRet = &H200) Then TextErr = TextErr + " -- IFCを受信しました <ステータス>": GoTo CheckEnd
    If (CheckRet = &H300) Then TextErr = TextErr + " -- SRQとIFCを受信しました <ステータス>"

CheckEnd:
    RetBuf = TextErr
    CheckRetCode = RetSts

End Function

Public Function DevidedString(Base_Str As String, Str_Cnt As Long) As String

'Base_Strの中から","で区切られたStr_Cnt番目の文字列を返します。
'Str_Cnt=1の時、先頭の文字列を返します。また、Str_Cnt番目の文字列が
'なかった場合,またStr_Cnt=0,Str_Cnt>100だった場合には ""を返します。
Dim StrLenPre(100) As Integer
Dim StrLenAft As Integer
Dim BaseLen As Integer
Dim Count As Integer

    If (Str_Cnt = 0) Or (Str_Cnt > 100) Or (Base_Str = "") Then
        DevidedString = ""
        Exit Function
    End If

    '渡された文字列の長さを取得します
    BaseLen = Len(Trim$(Base_Str))
    StrLenPre(1) = 0

    For Count = 1 To Str_Cnt
        '","のある位置を取得します
        StrLenAft = InStr(StrLenPre(Count) + 1, Base_Str, ",")
        If StrLenAft = 0 Then
            '指定された位置より後ろに","が見つからなかった場合
            If Count = Str_Cnt Then
                '区切られた最後の位置の場合
                StrLenAft = BaseLen + 1
            Else
                '区切られた数より、指定された位置(Str_Cnt)が
                '大きかった場合(結果として""を返します)
                StrLenAft = 1
            End If
            Exit For
        End If
        '次の検索開始位置の指定
        StrLenPre(Count + 1) = StrLenAft
    Next

    DevidedString = Trim$(Mid$(Base_Str, StrLenPre(Str_Cnt) + 1, StrLenAft - 1))

End Function
