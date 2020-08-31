Attribute VB_Name = "gpibModule"
'
'  2001.1.31 更新 白石 ATiC
'
Declare Function GpibOpen Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long
Declare Function GpibSetIfc Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long
Declare Function GpibSetRen Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long
Declare Function GpibSetConfig Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByVal szInfo As String) As Long
Declare Function GpibExecDeviceTrigger Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByRef lpAdrsTbl As Long) As Long
Declare Function GpibExecDeviceClear Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByRef lpAdrsTbl As Long) As Long
Declare Function GpibCheckSrq Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long
Declare Function GpibExecSpoll Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByRef lpAdrsTbl As Long, ByRef lpStbTbl As Long, ByRef lpStbAdrs As Long) As Long
Declare Function GpibReceive Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByRef lpAdrsTbl As Long, ByRef ulpLength As Long, ByVal vpBuffer As String) As Long
Declare Function GpibSend Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByRef lpAdrsTbl As Long, ByVal ulpLength As Long, ByVal vpBuffer As String) As Long
Declare Function GpibClose Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long
Declare Function GpibSetSrqEvent Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByVal lpOnSrqProc As Any, ByVal dwUser As Long) As Long
Declare Function GpibWaitSrqEvent Lib "GPC43042.DLL" (ByVal ulBoardNo As Long, ByVal ulTimeOut As Long) As Long
Declare Function GpibKillSrqEvent Lib "GPC43042.DLL" (ByVal ulBoardNo As Long) As Long

Public nBoardNo As Long              ' ボードアクセス番号
Public nInitBoardNo As Long          ' ボードアクセス番号（初期化指定番号）
Public DevAdrs As Long               ' 機器アドレス

Public DevAdrsTbl(1) As Long         ' 機器アドレステーブル

Public bCheckSrq As Boolean
Public Delim_S As String             ' 送信デリミタ
Public Delim_R As String             ' 送信デリミタ
Public SetPrm As String * 128        ' パラメータ
Public RecvBuffer As String * 32     ' 受信バッファ

Public OpenFlag As Boolean          ' Open:True , NOT Open:False
Public EventFlag As Long

Public Sub DsplyErrMessage(ByVal nErrCode As Integer)
    
    Dim strRetStatus As String
    
    Select Case nErrCode
        Case 2
            'nRet = MsgBox("エラーコード 2　　正常終了:ＥＯＩを検出して終了しました", (vbOKOnly + vbInformation), "エラーコード")
            strRetStatus = "コード 2　　正常終了:ＥＯＩを検出して終了しました"
        Case 1
            'nRet = MsgBox("エラーコード 1　　正常終了:指定された受信データ数に達して終了しました", (vbOKOnly + vbInformation), "エラーコード")
            strRetStatus = "コード 1　　正常終了:指定された受信データ数に達して終了しました"
        Case 0
            'nRet = MsgBox("エラーコード 0　　正常終了", (vbOKOnly + vbInformation), "エラーコード")
            strRetStatus = "コード 0　　正常終了"
        Case -1
            'nRet = MsgBox("エラーコード -1    ボードアクセス番号が違います", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -1    ボードアクセス番号が違います"
        Case -4
            'nRet = MsgBox("エラーコード -4    スレーブモードでは使用できません", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -4    スレーブモードでは使用できません"
        Case -5
            'nRet = MsgBox("エラーコード -5    マスタモードでは使用できません", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -5    マスタモードでは使用できません"
        Case -7
            'nRet = MsgBox("エラーコード -7    バスコマンドの送出に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -7    バスコマンドの送出に失敗しました"
        Case -12
            'nRet = MsgBox("エラーコード -12   データ受信に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -12   データ受信に失敗しました"
        Case -13
            'nRet = MsgBox("エラーコード -13   データ送信に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -13   データ送信に失敗しました"
        Case -14
            'nRet = MsgBox("エラーコード -14   タイムアウトが発生しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -14   タイムアウトが発生しました"
        Case -16
            'nRet = MsgBox("エラーコード -16   ＩＦＣ受信による強制終了", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -16   ＩＦＣ受信による強制終了"
        Case -20
            'nRet = MsgBox("エラーコード -20   現在、バスが占有状態となっています", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -20   現在、バスが占有状態となっています"
        Case -40
            'nRet = MsgBox("エラーコード -40   コールバックイベントの登録に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -40   コールバックイベントの登録に失敗しました"
        Case -41
            'nRet = MsgBox("エラーコード -41   コールバックイベントの登録解除に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -41   コールバックイベントの登録解除に失敗しました"
        Case -42
            'nRet = MsgBox("エラーコード -42   イベントオブジェクトが有効（シグナル状態）です", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -42   イベントオブジェクトが有効（シグナル状態）です"
        Case -43
            'nRet = MsgBox("エラーコード -43   指定した時間内にイベントオブジェクトが有効にならなかったためタイムアウトしました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -43   指定した時間内にイベントオブジェクトが有効にならなかったためタイムアウトしました"
        Case -44
            'nRet = MsgBox("エラーコード -44   イベントオブジェクト待ちでエラーが発生しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -44   イベントオブジェクト待ちでエラーが発生しました"
        Case -45
            'nRet = MsgBox("エラーコード -45   コールバックイベントが登録されていません", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -45   コールバックイベントが登録されていません"
        Case -995
            'nRet = MsgBox("エラーコード -995  ボードの使用終了ができませんでした", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -995  ボードの使用終了ができませんでした"
        Case -996
            'nRet = MsgBox("エラーコード -996  ドライバ側のページ確保ができませんでした", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -996  ドライバ側のページ確保ができませんでした"
        Case -997
            'nRet = MsgBox("エラーコード -997  タイマ設定に失敗しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -997  タイマ設定に失敗しました"
        Case -998
            'nRet = MsgBox("エラーコード -998  割り込みが使用できません", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -998  割り込みが使用できません"
        Case -999
            'nRet = MsgBox("エラーコード -999　ボードが存在しない、またはボードのI/Oポートに異常があります", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "コード -999　ボードが存在しない、またはボードのI/Oポートに異常があります"
        Case Else
            'nRet = MsgBox("予想外のエラーが発生しました", (vbOKOnly + vbCritical), "エラーコード")
            strRetStatus = "予想外のエラーが発生しました"
    End Select

    計測Form.StatusBar1.Panels(1).Text = strRetStatus

End Sub

' -----------------------------------------------------------------------
'  コールバックモジュール
' -----------------------------------------------------------------------

Sub lpOnSrqProc(ByVal dwBoardNo As Long, ByVal dwUser As Long)

    ' 割り込みイベント処理を行うコールバック関数です
    ' VB5.0以上で使用できます
    ' ここにイベントが発生した時の処理を書いてください
    On Error Resume Next
    
    Dim nStbTbl(4) As Long
    Dim nStbAdrs(4) As Long
    Dim szDisp As String
    
    ' 機器アドレステーブル設定
    DevAdrsTbl(0) = DevAdrs
    DevAdrsTbl(1) = -1
    
    ' シリアルポーリング
    nRet = GpibExecSpoll(nBoardNo, DevAdrsTbl(0), nStbTbl(0), nStbAdrs(0))
    
    If nRet Then
        Call DsplyErrMessage(nRet)
    Else
        CtrlSpollForm.LblStat.Caption = Hex(nStbTbl(0))
        bCheckSrq = False
    End If
    
    szDisp = "ステータス値 : " & Hex(nStbTbl(0)) & " HEX"
    nRet = MsgBox(szDisp, (vbOKOnly + vbInformation), "ステータス")

End Sub

Function GPinit() As Boolean
    Dim nRet As Long
    Dim UserData As Long

    
    ' ボード番号取得
    nBoardNo = GetIni("GP-IB設定", "ボード番号", CurrentDIR & "計測設定.ini")

    ' 機器アドレス取得
    DevAdrs = GetIni("GP-IB設定", "機器アドレス", CurrentDIR & "計測設定.ini")
    
    '送信デリミタコード
    Delim_S = GetIni("GP-IB設定", "送信デリミタコード", CurrentDIR & "計測設定.ini") '"CRLF+EOI"
    
    '受信デリミタコード
    Delim_R = GetIni("GP-IB設定", "受信デリミタコード", CurrentDIR & "計測設定.ini") '"CRLF+EOI"
    
    ' ボード初期化
    
    nRet = GpibOpen(nBoardNo)
    If nRet Then
        Call DsplyErrMessage(nRet)
        GoTo ExitInit
    End If
    
    OpenFlag = True

    ' デリミタ設定
    SetPrm = "/SRT=50 /TMO=50 /SDELIM=" + Delim_S + " /RDELIM=" + Delim_R ' INIファイルから取得する
    nRet = GpibSetConfig(nBoardNo, SetPrm)
    If nRet Then
        Call DsplyErrMessage(nRet)
        GoTo ExitInit
    End If
  
    ' ＩＦＣ送出
    nRet = GpibSetIfc(nBoardNo)
    If nRet Then
        Call DsplyErrMessage(nRet)
        Unload CtrlInitForm
        GoTo ExitInit
    End If
    
    ' ＲＥＮ設定
    nRet = GpibSetRen(nBoardNo)
    Call DsplyErrMessage(nRet)
    
    ' イベントの設定は、VB5.0以降でのみ行えます。
    ' イベントを使用する場合は、VB5.0以降でコメントをはずしてご使用ください。
    ' イベント設定
'    UserData = &H55
'    nRet = GpibSetSrqEvent(nBoardNo, AddressOf lpOnSrqProc, UserData)
'    EventFlag = 1
    
    GPinit = True
    Exit Function
    
ExitInit:
    GPinit = False
    
End Function

Sub GPclose()
    OpenFlag = False
    nRet = GpibClose(nBoardNo)
End Sub



