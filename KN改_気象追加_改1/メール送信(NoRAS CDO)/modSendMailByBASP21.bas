Attribute VB_Name = "modSendMailByBASP21"
'*******************************************************************************
'   Eﾒｰﾙ送信機能 ※BSMTP.dll(BASP21),UNLHA32.dll必須
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
Option Explicit
'**********↓↓↓実際に送信を行なう段階でこの値を「1」に変更して下さい↓↓↓
#Const cnsSW_TEST = 0       ' テスト中(=0)
'#Const cnsSW_TEST = 1       ' 本番(=1)
'**********↑↑↑実際に送信を行なう段階でこの値を「1」に変更して下さい↑↑↑
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4
Private Const g_cnsNG = "NG"
Private Const g_cnsOK = "OK"
Private Const g_cnsYen = "\"
Private Const g_cnsLZH = ".lzh"
Private Const g_cnsERRMSG1 = "が正しくありません。"
Private Const g_cnsCNT1 = 3     ' 格納ﾃｰﾌﾞﾙ1次元目の要素数(固定)
Private Const MAX_PATH = 260
' ﾒｰﾙ送信API(BASP21)
Private Declare Function SendMail Lib "BSMTP.dll" _
    (szServer As String, szTo As String, szFrom As String, _
     szSubject As String, szBody As String, szFile As String) As String
' LHA圧縮を操作するAPI(UNLHA32)
Private Declare Function Unlha Lib "UNLHA32.dll" _
    (ByVal lhWnd As Long, ByVal szCmdLine As String, _
     ByVal szOutPut As String, ByVal wSize As Long) As Long
' ﾀﾞｲｱﾙｱｯﾌﾟｴﾝﾄﾘｰを指定して接続(IE4以上必須)
Private Declare Function InternetDial Lib "WININET.dll" _
    (ByVal hwndParent As Long, ByVal lpszConnectoid As String, _
     ByVal dwFlags As Long, lpdwConnection As Long, _
     ByVal dwReserved As Long) As Long
' ﾀﾞｲｱﾙｱｯﾌﾟIDを指定して切断
Private Declare Function InternetHangUp Lib "WININET.dll" _
    (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
' ｳｨﾝﾄﾞｩﾊﾝﾄﾞﾙを返す
Private Declare Function FindWindow Lib "USER32.dll" _
    Alias "FindWindowA" (ByVal lpClassName As Any, _
    ByVal lpWindowName As Any) As Long
' Sleep
Private Declare Sub Sleep Lib "KERNEL32.dll" _
    (ByVal dwMilliseconds As Long)
' SYSTEMﾃﾞｨﾚｸﾄﾘ名取得API
Private Declare Function GetSystemDirectory Lib "KERNEL32.dll" _
    Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
' WindowsのTEMPﾌｫﾙﾀﾞ取得
Private Declare Function GetTempPath Lib "KERNEL32.dll" _
    Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'*******************************************************************************
' Eﾒｰﾙ送信機能(BSMTP.dll必須)
'*******************************************************************************
' [引数]
'   strDialUp   : ﾀﾞｲｱﾙｱｯﾌﾟ登録名(ﾀﾞｲｱﾙｱｯﾌﾟしない時はﾌﾞﾗﾝｸ)
'   strDomain   : ﾄﾞﾒｲﾝ名(xxxx.co.jp等)
'   strSMTP     : SMTPｻｰﾊﾞ名(smtp.xxxx.co.jp,mail.xxxx.co.jp等)
'   strPort     : 通常は｢25｣,ﾌﾞﾗﾝｸの場合は｢25｣
'   strTimeOut  : ｢60｣位が適当,ﾌﾞﾗﾝｸの場合は｢60｣
'   strFromName : 送信元名称
'   strFromAddr : 送信元ｱﾄﾞﾚｽ
'   vntToName   : 宛先名称(複数の場合は配列をｾｯﾄする)
'   vntToAddr   : 宛先ｱﾄﾞﾚｽ(複数の場合は配列をｾｯﾄする,配列要素数は宛先名称と一致させる)
'   vntCCName   : CC宛先名称(複数の場合は配列をｾｯﾄする)
'   vntCCAddr   : CC宛先ｱﾄﾞﾚｽ(複数の場合は配列をｾｯﾄする,配列要素数はCC名と一致させる)
'   vntBCCName  : BCC宛先名称(複数の場合は配列をｾｯﾄする)
'   vntBCCAddr  : BCC宛先ｱﾄﾞﾚｽ(複数の場合は配列をｾｯﾄする,配列要素数はBCC名と一致させる)
'   swOwnerBCC  : Trueの場合､送信元ｱﾄﾞﾚｽをBCCに加える
'   strSubj     : 件名
'   strMessage  : 本文(署名も付加してｾｯﾄ)
'   strCaption  : 親ｳｨﾝﾄﾞｳのCaption
'   vntFileName : ﾌﾙﾊﾟｽ添付ﾌｧｲﾙ名(複数の場合は配列をｾｯﾄする) ※ない場合はﾌﾞﾗﾝｸ
'   strLzhFile  : 上記添付ﾌｧｲﾙを圧縮する場合はその圧縮ﾌｧｲﾙ名(ﾊﾟｽ名不要)
'   intDelMode  : 圧縮時の削除方法(0=削除なし, 1=圧縮ﾌｧｲﾙを削除, 2=元ﾌｧｲﾙを削除)
' [戻り値]
'   "OK"=成功, それ以外はｴﾗｰﾒｯｾｰｼﾞ
'*******************************************************************************
Public Function SendMailByBASP21(strDialUp As String, _
                                 strDomain As String, _
                                 strSMTP As String, _
                                 strPort As String, _
                                 strTimeOut As String, _
                                 strFromName As String, _
                                 strFromAddr As String, _
                                 vntToName As Variant, _
                                 vntToAddr As Variant, _
                                 vntCCName As Variant, _
                                 vntCCAddr As Variant, _
                                 vntBCCName As Variant, _
                                 vntBCCAddr As Variant, _
                                 swOwnerBCC As Boolean, _
                                 strSubj As String, _
                                 strMessage As String, _
                                 Optional strCaption As String, _
                                 Optional vntFileName As Variant, _
                                 Optional strLzhFile As String, _
                                 Optional intDelMode As Integer) As String
    Dim xlAPP As Application
    Dim strDIAL_ENTRY As String     ' ﾀﾞｲｱﾙｱｯﾌﾟｴﾝﾄﾘｰ
    Dim strSV_Name As String        ' ﾄﾞﾒｲﾝ/SMTP:ﾎﾟｰﾄ:ﾀｲﾑｱｳﾄ
    Dim strMailFrom As String       ' 送信元登録
    Dim strMailto As String         ' 送信先登録
    Dim strTable() As String        ' 配列値格納ﾃｰﾌﾞﾙ
    Dim MAX2 As Integer             ' ﾃｰﾌﾞﾙに格納した要素数(2次元目最大値)
    Dim CNT2() As Integer           ' ﾃｰﾌﾞﾙに格納した要素数(2次元目各要素)
    Dim vntName As Variant          ' 宛先名Work
    Dim vntAddr As Variant          ' ｱﾄﾞﾚｽWork
    Dim strPathName As String       ' 添付ﾌｧｲﾙのﾌｫﾙﾀﾞ名
    Dim strFileName As String       ' 添付ﾌｧｲﾙ
    Dim swLine As Byte              ' ﾀﾞｲｱﾙ接続ｽｲｯﾁ
    Dim hWnd As Long                ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ
    Dim lngConnID As Long           ' ｺﾈｸｼｮﾝｺｰﾄﾞ
    Dim IX As Long                  ' ﾃｰﾌﾞﾙIndex
    Dim IX1 As Long                 ' ﾃｰﾌﾞﾙIndex
    Dim IX2 As Long                 ' ﾃｰﾌﾞﾙIndex
    Dim IX3 As Long                 ' ﾃｰﾌﾞﾙIndex
    Dim lngRet As Long              ' ﾘﾀｰﾝｺｰﾄﾞ
    Dim strRC As String             ' BASP21戻り値
    Dim strMSG As String            ' ﾒｯｾｰｼﾞ
    Dim vntMSG As Variant           ' ﾒｯｾｰｼﾞWork
    Dim strName As String           ' Work
    Dim strAddr As String           ' Work

    SendMailByBASP21 = g_cnsNG
    Set xlAPP = Application
    
    ' BSMTP.dllの存在確認
    If Dir(FP_GET_SYSTEM_PATH & "BSMTP.dll", vbNormal) = "" Then
        SendMailByBASP21 = _
            "送信コンポーネント｢BSMTP.dll｣がインストールされていません。"
        Exit Function
    End If
    
'-------------------------------------------------------------------------------
' ■準備処理(引き渡しパラメータの作成)
    
    ' ﾄﾞﾒｲﾝ/SMTP:ﾎﾟｰﾄ:ﾀｲﾑｱｳﾄ
    If Trim$(strPort) = "" Then strPort = "25"
    If Trim$(strTimeOut) = "" Then strTimeOut = "60"
    strSV_Name = Trim$(strDomain) & "/" & _
                 Trim$(strSMTP) & ":" & _
                 Trim$(strPort) & ":" & _
                 Trim$(strTimeOut)
    
    ' Variant項目未使用の場合の対応
    If IsError(vntToName) Then vntToName = ""
    If IsError(vntToAddr) Then vntToAddr = ""
    If IsError(vntCCName) Then vntCCName = ""
    If IsError(vntCCAddr) Then vntCCAddr = ""
    If IsError(vntBCCName) Then vntBCCName = ""
    If IsError(vntBCCAddr) Then vntBCCAddr = ""
    If IsError(vntFileName) Then vntFileName = ""
                 
    ' 送信元登録
    If Trim$(strFromAddr) = "" Then
        SendMailByBASP21 = "送信元のメールアドレスがありません"
        Exit Function
    End If
    If Trim$(strFromName) = "" Then
        strMailFrom = Trim$(strFromAddr)
    Else
        strMailFrom = Trim$(strFromName) & _
            "<" & Trim$(strFromAddr) & ">"
    End If
    
    ' 配列で引き渡される可能性がある項目を全て別ﾃｰﾌﾞﾙに格納し直す
    ' (以後は全て配列の文字列変数として処理できる)
    MAX2 = 0
    ReDim strTable(g_cnsCNT1, MAX2)
    ReDim CNT2(g_cnsCNT1)
    vntMSG = Array("宛先", "CC宛先", "BCC宛先", "添付ファイル名")
    For IX1 = 0 To g_cnsCNT1
        Select Case IX1
            Case 0: vntName = vntToName:   vntAddr = vntToAddr      ' 宛先
            Case 1: vntName = vntCCName:   vntAddr = vntCCAddr      ' CC
            Case 2: vntName = vntBCCName:  vntAddr = vntBCCAddr     ' BCC
            Case 3: vntName = vntFileName: vntAddr = vntFileName    ' 添付ﾌｧｲﾙ
        End Select
        IX3 = 0
        If IsArray(vntAddr) = True Then
            ' 格納ﾃｰﾌﾞﾙに配列を格納
            For IX2 = LBound(vntAddr) To UBound(vntAddr)
                On Error GoTo MakeArray_ARRAY2
                strAddr = Trim$(vntAddr(IX2))
                If ((IX1 < g_cnsCNT1) And (IX2 <= UBound(vntName))) Then
                    On Error GoTo MakeArray_ARRAY3
                    strName = Trim$(vntName(IX2))
                Else
                    strName = ""
                End If
                GoSub MakeArray_SUB
            Next IX2
        Else
            If IX1 < g_cnsCNT1 Then
                strName = Trim$(vntName)
            Else
                strName = ""
            End If
            strAddr = Trim$(vntAddr)
            GoSub MakeArray_SUB
        End If
        CNT2(IX1) = IX3
    Next IX1
    If CNT2(0) < 1 Then
        SendMailByBASP21 = "宛先のメールアドレスがありません"
        Exit Function
    End If
    
    ' 送信者をBCCに追加指定の処理(swOwnerBCC指定の場合)
    If swOwnerBCC = True Then
        CNT2(2) = CNT2(2) + 1
        If CNT2(2) > MAX2 Then
            MAX2 = CNT2(2)
            ReDim Preserve strTable(g_cnsCNT1, MAX2)
        End If
        If strFromName <> "" Then
            strTable(2, CNT2(2)) = strFromName & "<" & strFromAddr & ">"
        Else
            strTable(2, CNT2(2)) = strFromAddr
        End If
    End If
    
    ' 送信先登録(宛先,CC,BCCをTab区切りﾃｷｽﾄにする)
    strMailto = ""
    For IX1 = 0 To 2
        If CNT2(IX1) >= 1 Then
            Select Case IX1
                Case 1: strMailto = strMailto & vbTab & "cc"
                Case 2: strMailto = strMailto & vbTab & "bcc"
            End Select
            IX = 1
            Do While IX <= CNT2(IX1)
                ' 2件目以降はTab区切りでｾｯﾄ
                If strMailto = "" Then
                    strMailto = strTable(IX1, IX)
                Else
                    strMailto = strMailto & vbTab & strTable(IX1, IX)
                End If
                IX = IX + 1
            Loop
        End If
    Next IX1
    
    ' 添付ﾌｧｲﾙ処理
    strFileName = ""
    If Trim$(strLzhFile) <> "" Then
        ' 圧縮ﾌｧｲﾙが指定されている場合は圧縮ﾌｧｲﾙを添付ﾌｧｲﾙに指定(単一ﾌｧｲﾙ)
        strMSG = FP_ArchiveByUNLHA32(strLzhFile, vntFileName, strCaption)
        If strMSG <> g_cnsOK Then
            SendMailByBASP21 = "圧縮ファイルの作成に失敗しました。" & vbCr & _
                strMSG
            Exit Function
        End If
        strFileName = strLzhFile
    Else
        ' 圧縮ﾌｧｲﾙが指定されていない場合はTab区切りﾃｷｽﾄにする
        IX1 = g_cnsCNT1
        IX = 1
        Do While IX <= CNT2(IX1)
            If strFileName = "" Then
                strFileName = strTable(IX1, IX)
            Else
                strFileName = strFileName & vbTab & strTable(IX1, IX)
            End If
            IX = IX + 1
        Loop
    End If
    
'-------------------------------------------------------------------------------
' ■送信処理
    
    ' ﾀﾞｲｱﾙｱｯﾌﾟ接続(ｴﾝﾄﾘ名が指定されている場合のみ)
    strDIAL_ENTRY = Trim$(strDialUp)
    If strDIAL_ENTRY <> "" Then
        strMSG = ""
        swLine = 0
        xlAPP.StatusBar = "｢" & strDIAL_ENTRY & "｣に接続中です．．．．"
        ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙを取得
        hWnd = FP_GET_HWND(strCaption)
        lngConnID = 0
        ' ﾘﾓｰﾄ接続を起動
        lngRet = InternetDial(hWnd, strDIAL_ENTRY, _
            INTERNET_AUTODIAL_FORCE_UNATTENDED, lngConnID, 0&)
        If ((lngRet <> 0) And (lngRet <> 633)) Then
            strMSG = "｢" & strDIAL_ENTRY & "｣への接続に失敗しました。"
            Select Case lngRet
                Case 623: strMSG = strMSG & vbCr & "　(ﾀﾞｲｱﾙｴﾝﾄﾘｰ名が不存在)"
                Case 668: strMSG = strMSG & vbCr & "　(ﾊﾟｽﾜｰﾄﾞが未登録)"
                Case Else: strMSG = strMSG & vbCr & _
                    "　(その他ｴﾗｰ : " & CStr(lngRet) & " )"
            End Select
            SendMailByBASP21 = strMSG
            Exit Function
        End If
        swLine = 1
    End If
    
    ' BASP21(BSMTP.dll)実行
    xlAPP.StatusBar = "メールを送信中です．．．．"
    On Error GoTo BASP_ERROR
#If cnsSW_TEST = 1 Then
    ' 本番
    strRC = SendMail(strSV_Name, strMailto, strMailFrom, strSubj, _
        strMessage, strFileName)
#Else
    ' テスト(引数表示のみ)
    MsgBox "・ﾄﾞﾒｲﾝ/SMTP:ﾎﾟｰﾄ:ﾀｲﾑｱｳﾄ = " & strSV_Name & vbCr & _
           "・宛先 = " & strMailto & vbCr & _
           "・差出人 = " & strMailFrom & vbCr & _
           "・件名 = " & strSubj & vbCr & _
           "・添付 = " & strFileName & vbCr & vbCr & _
           "※これはテスト用の確認メッセージです。" & vbCr & _
           "　本番に切り替えるには、modSendMailByBASP21_2の最初にある" & vbCr & _
           "　コンパイルスイッチ「cnsSW_TEST」の値を「1」に変更して保存して下さい。"
#End If
    
    ' ﾀﾞｲｱﾙｺﾈｸｼｮﾝを切断
    If swLine = 1 Then
        ' 回線を切断
        xlAPP.StatusBar = "｢" & strDIAL_ENTRY & "｣を切断中です．．．．"
        InternetHangUp lngConnID, 0&
        AppActivate xlAPP.Caption       ' Excelをｱｸﾃｨﾌﾞにする
        swLine = 0
    End If
    
    If strRC <> "" Then
        SendMailByBASP21 = strRC & vbCr & vbCr & _
            "サーバーに接続できないか、切断されました。"
        xlAPP.StatusBar = False
        Exit Function
    End If
    
'-------------------------------------------------------------------------------
' ■終了処理(圧縮ﾌｧｲﾙ指定時の事後削除処理)
    
    ' 圧縮ﾌｧｲﾙを作成した場合は削除するか判定する(送信正常時のみ)
    If strLzhFile <> "" Then
        xlAPP.DisplayAlerts = False
        Select Case intDelMode
            Case 1
                ' 圧縮ﾌｧｲﾙを削除する
                Kill strLzhFile
            Case 2
                ' 元ﾌｧｲﾙを削除する
                If IsArray(vntFileName) = True Then
                    ' 配列指定時は順次削除
                    vntAddr = vntFileName
                    For IX2 = LBound(vntAddr) To UBound(vntAddr)
                        On Error GoTo MakeArray_ARRAY2
                        strAddr = Trim$(vntAddr(IX2))
                        On Error Resume Next
                        Kill strAddr
                    Next IX2
                    On Error GoTo 0
                Else
                    ' 単一ﾌｧｲﾙ指定
                    strFileName = Trim$(vntFileName)
                    On Error Resume Next
                    Kill strFileName
                    On Error GoTo 0
                End If
        End Select
        xlAPP.DisplayAlerts = True
    End If
    
    SendMailByBASP21 = g_cnsOK
    AppActivate xlAPP.Caption           ' Excelをｱｸﾃｨﾌﾞにする
    xlAPP.StatusBar = False
    Exit Function

'-------------------------------------------------------------------------------
' 1次元参照でｴﾗｰの場合は2次元として処理(ｾﾙ範囲格納対応)
MakeArray_ARRAY2:
    On Error GoTo MakeArray_ERROR
    strAddr = Trim$(vntAddr(IX2, 1))
    Resume Next

'-------------------------------------------------------------------------------
' 1次元参照でｴﾗｰの場合は2次元として処理(ｾﾙ範囲格納対応)
MakeArray_ARRAY3:
    On Error GoTo MakeArray_ERROR
    strName = Trim$(vntName(IX2, 1))
    Resume Next

'-------------------------------------------------------------------------------
' 格納ﾃｰﾌﾞﾙにｾｯﾄする
MakeArray_SUB:
    If strAddr <> "" Then
        IX3 = IX3 + 1
        If IX3 > MAX2 Then
            ' 最大値で格納ﾃｰﾌﾞﾙの要素数を変更
            MAX2 = IX3
            ReDim Preserve strTable(g_cnsCNT1, MAX2)
        End If
        If strName <> "" Then
            strTable(IX1, IX3) = strName & "<" & strAddr & ">"
        Else
            strTable(IX1, IX3) = strAddr
        End If
        ' 宛先に送信者ｱﾄﾞﾚｽがある場合はBCC付加しない
        If strAddr = strFromAddr Then swOwnerBCC = False
    End If
    Return

'-------------------------------------------------------------------------------
' ｴﾗｰ処理
MakeArray_ERROR:
    SendMailByBASP21 = "パラメータ登録処理に失敗しました。(" & _
        vntMSG(IX1) & ")" & vbCr & "  (" & Err.Description & ")"
    xlAPP.StatusBar = False
    Exit Function

'-------------------------------------------------------------------------------
' BASP21実行時ｴﾗｰ
BASP_ERROR:
    strRC = "メール送信コンポーネント｢BASP21｣が実行できません。" & _
        vbCr & Err.Number & " " & Err.Description
    Resume Next
    
End Function

'*******************************************************************************
' ﾌｧｲﾙ圧縮機能(UNLHA32.dll必須)
'*******************************************************************************
' [引数]
'   strTarget   : 圧縮後のﾌｧｲﾙ名
'   vntSource   : 圧縮対象のﾌｧｲﾙ名(複数の場合は配列をｾｯﾄする)
'   strCaption  : 親ｳｨﾝﾄﾞｳのCaption
' [戻り値]
'   "OK"=成功, それ以外はｴﾗｰﾒｯｾｰｼﾞ
'*******************************************************************************
Public Function FP_ArchiveByUNLHA32(strTarget As String, _
                                    vntSource As Variant, _
                                    Optional strCaption As String) As String
    Dim xlAPP As Application
    Dim strFileName As String       ' ﾌｧｲﾙ名(work)
    Dim strPathName As String       ' 圧縮ﾌｧｲﾙのﾌｫﾙﾀﾞ
    Dim strExeName As String        ' 自動解凍圧縮ﾌｧｲﾙ
    Dim strCommand As String        ' UNLHAｺﾏﾝﾄﾞﾗｲﾝ
    Dim strBuffer As String         ' Work
    Dim strEXT As String            ' 拡張子
    Dim IX As Long                  ' ﾃｰﾌﾞﾙIndex
    Dim hWnd As Long                ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ
    Dim strMSG As String            ' UNLHA32ｴﾗｰﾒｯｾｰｼﾞ
    Dim cntSource As Long           ' 入力ﾌｧｲﾙ数
    
    FP_ArchiveByUNLHA32 = g_cnsNG
    Set xlAPP = Application
    xlAPP.StatusBar = "圧縮ファイル作成中．．．．"
    
    ' UNLHA32.dllの存在確認
    If Dir(FP_GET_SYSTEM_PATH & "UNLHA32.dll", vbNormal) = "" Then
        FP_ArchiveByUNLHA32 = _
            "圧縮コンポーネント｢UNLHA32｣がインストールされていません。"
        Exit Function
    End If
    
    ' 出力ﾌｧｲﾙにﾊﾟｽがない場合はTEMPﾌｫﾙﾀﾞに出力
    strFileName = Trim$(strTarget)
    If ((Left$(strFileName, 2) <> "\\") And _
        (Mid$(strFileName, 2, 2) <> ":\")) Then
        ' TEMPﾌｫﾙﾀﾞを受領
        strPathName = FP_GET_TEMP_PATH
        strTarget = strPathName & strFileName
    Else
        strTarget = strFileName
        IX = Len(strFileName)
        Do While IX > 1
            If Mid$(strFileName, IX, 1) = g_cnsYen Then Exit Do
            IX = IX - 1
        Loop
        strPathName = Left$(strFileName, IX)
    End If
    
    ' 拡張子の判定
    strEXT = StrConv(Right$(strTarget, 3), vbUpperCase)
    If ((strEXT <> "LZH") And (strEXT <> "EXE")) Then
        If Mid$(strTarget, Len(strTarget) - 3, 1) <> "." Then
            strTarget = strTarget & g_cnsLZH
        Else
            strFileName = Left$(strTarget, Len(strTarget) - 4)
            strTarget = strFileName & g_cnsLZH
        End If
    ElseIf strEXT = "EXE" Then
        strExeName = strTarget
        strFileName = Left$(strTarget, Len(strTarget) - 4)
        strTarget = strFileName & g_cnsLZH
    End If
    
    ' 同名ﾌｧｲﾙが存在する場合は削除
    If Dir(strTarget, vbNormal) <> "" Then Kill strTarget
    
    ' UNLHAのｺﾏﾝﾄﾞﾗｲﾝを編集
    strCommand = "a """ & strTarget & """"
    If IsArray(vntSource) = True Then
        For IX = LBound(vntSource) To UBound(vntSource)
            On Error GoTo UnLha_ARRAY
            strFileName = Trim$(vntSource(IX))
            On Error GoTo 0
            If strFileName <> "" Then
                If GetAttr(strFileName) And vbDirectory Then
                    FP_ArchiveByUNLHA32 = _
                        "複数指定ではフォルダは指定できません。"
                    GoTo UnLha_EXIT
                Else
                    strCommand = strCommand & " """ & strFileName & """"
                End If
                cntSource = cntSource + 1
            End If
        Next IX
    ElseIf IsError(vntSource) <> True Then
        strFileName = Trim$(vntSource)
        If strFileName <> "" Then
            If GetAttr(strFileName) And vbDirectory Then
                ' ﾌｫﾙﾀﾞ指定の場合は配下全てを格納
                strCommand = strCommand & " -d1 """ & _
                    Left(strFileName, Len(strFileName) - _
                        Len(Dir(strFileName, vbDirectory))) & "\"" " & _
                    Dir(strFileName, vbDirectory)
            Else
                strCommand = strCommand & " """ & strFileName & """"
            End If
            cntSource = cntSource + 1
        End If
    End If
    
    ' 有効な入力ﾌｧｲﾙがない場合は無視
    If cntSource < 1 Then
        strTarget = ""
        FP_ArchiveByUNLHA32 = g_cnsOK
        Exit Function
    End If
    
    On Error GoTo UnLha_ERROR
    ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙを取得
    hWnd = FP_GET_HWND(strCaption)
    ' ｺﾏﾝﾄﾞﾗｲﾝに従ってUNLHAを操作
    strBuffer = String(256, Chr$(0))
    If Unlha(hWnd, strCommand, strBuffer, Len(strBuffer)) = 0& Then
        If strEXT = "EXE" Then
            ' EXE形式指定の場合は自動解凍書庫に変換
            strCommand = "s -gw2 """ & strTarget & """ """ & strPathName & """"
            strBuffer = String(256, Chr$(0))
            If Unlha(hWnd, strCommand, strBuffer, Len(strBuffer)) = 0& Then
                Kill strTarget
                strTarget = strExeName
                FP_ArchiveByUNLHA32 = g_cnsOK
            Else
                FP_ArchiveByUNLHA32 = _
                    Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
            End If
        Else
            FP_ArchiveByUNLHA32 = g_cnsOK
        End If
    Else
        FP_ArchiveByUNLHA32 = Left$(strBuffer, InStr(1, strBuffer, Chr$(0)) - 1)
    End If
    GoTo UnLha_EXIT

'-------------------------------------------------------------------------------
' 配列操作ｴﾗｰ対応(2次元配列の場合は再配置して戻る)
UnLha_ARRAY:
    On Error GoTo UnLha_ERROR2
    strFileName = Trim$(vntSource(IX, 1))
    Resume Next
    
'-------------------------------------------------------------------------------
' UNLHA32実行時ｴﾗｰ
UnLha_ERROR:
    FP_ArchiveByUNLHA32 = "圧縮コンポーネント｢UNLHA32｣が実行できません。" & _
        vbCr & Err.Number & " " & Err.Description
    GoTo UnLha_EXIT

'-------------------------------------------------------------------------------
' 配列操作時ｴﾗｰ
UnLha_ERROR2:
    FP_ArchiveByUNLHA32 = "入力ファイル指定が正しくありません。(UNLHA32)" & _
        vbCr & Err.Number & " " & Err.Description

'-------------------------------------------------------------------------------
' UNLHA32処理終了
UnLha_EXIT:
    On Error Resume Next
    AppActivate xlAPP.Caption
End Function

'*******************************************************************************
' WindowsのSYSTEMフォルダ取得
'*******************************************************************************
' [戻り値] SYSTEMﾌｫﾙﾀﾞ(ｴﾗｰ無視)
'*******************************************************************************
Private Function FP_GET_SYSTEM_PATH() As String
    Dim strBuffer As String
    Dim strPathName As String
    
    ' Bufferを確保
    strBuffer = String(MAX_PATH, Chr(0))
    ' SYSTEMﾃﾞｨﾚｸﾄﾘ名取得
    Call GetSystemDirectory(strBuffer, MAX_PATH)
    ' Null文字の手前までを有効として表示(ｶｯｺ内はﾛﾝｸﾞﾌｧｲﾙ名変換後)
    strPathName = Left$(strBuffer, InStr(1, strBuffer, Chr(0)) - 1)
    If Right$(strPathName, 1) <> g_cnsYen Then strPathName = strPathName & g_cnsYen
    FP_GET_SYSTEM_PATH = strPathName
End Function

'*******************************************************************************
' WindowsのTEMPフォルダ取得
'*******************************************************************************
' [戻り値] TEMPﾌｫﾙﾀﾞ(ｴﾗｰ無視)
'*******************************************************************************
Private Function FP_GET_TEMP_PATH() As String
    Dim strBuffer As String
    Dim strPathName As String
    
    ' Bufferを確保
    strBuffer = String(MAX_PATH, Chr(0))
    ' SYSTEMﾃﾞｨﾚｸﾄﾘ名取得
    Call GetTempPath(MAX_PATH, strBuffer)
    ' Null文字の手前までを有効として表示(ｶｯｺ内はﾛﾝｸﾞﾌｧｲﾙ名変換後)
    strPathName = Left$(strBuffer, InStr(1, strBuffer, Chr(0)) - 1)
    If Right$(strPathName, 1) <> g_cnsYen Then strPathName = strPathName & g_cnsYen
    FP_GET_TEMP_PATH = strPathName
End Function

'*******************************************************************************
' ウィンドウハンドルの取得
'*******************************************************************************
' [引数]
'   strCaption  : ｳｨﾝﾄﾞｳのCaption(ｸﾗｽは自動判断)
' [戻り値]
'   hWnd        : ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ値(失敗はｾﾞﾛ)
'*******************************************************************************
Private Function FP_GET_HWND(strCaption As String) As Long
    Dim strClassName As String
    
    strClassName = "XLMAIN"
    Select Case strCaption
        Case "": strCaption = Application.Caption
        Case Application.Caption
        Case Else
            ' UserFormの場合
            If Val(Application.Version) <= 8 Then
                strClassName = "ThunderXFrame"      ' Excel97
            Else
                strClassName = "ThunderDFrame"      ' Excel2000以降
            End If
    End Select
    On Error GoTo GET_HWND_ERR
    FP_GET_HWND = FindWindow(strClassName, strCaption)
    Exit Function

'-------------------------------------------------------------------------------
GET_HWND_ERR:
    FP_GET_HWND = 0&
End Function

'-----------------------------<< End of Source >>-------------------------------
