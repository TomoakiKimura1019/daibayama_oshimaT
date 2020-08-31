Attribute VB_Name = "modCDOmail"
'*******************************************************************************
'   CDOでメールを送信する   ※参照設定版
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
'   [参照設定]
'   ・Microsoft CDO for Windows 2000 Library
'     (or Microsoft CDO for Exchange 2000 Library)
'*******************************************************************************
Option Explicit

'**********↓↓↓実際に送信を行なう段階でこの値を「1」に変更して下さい↓↓↓
'#Const cnsSW_TEST = 0       ' テスト中(=0)
#Const cnsSW_TEST = 1       ' 本番(=1)
'**********↑↑↑実際に送信を行なう段階でこの値を「1」に変更して下さい↑↑↑
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Const INTERNET_AUTODIAL_FAILIFSECURITYCHECK = 4
Private Const g_cnsNG = "NG"
Private Const g_cnsOK = "OK"
Private Const g_cnsYen = "\"
Private Const g_cnsERRMSG1 = "が正しくありません。"
Private Const g_cnsCNT1 = 3     ' 格納ﾃｰﾌﾞﾙ1次元目の要素数(固定)
Private Const MAX_PATH = 260

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

'*******************************************************************************
' メール送信(CDO)  ※参照設定版
'*******************************************************************************
' [引数]
'  ①MailSmtpServer : SMTPサーバ名(又はIPアドレス)
'  ②MailFrom       : 送信元アドレス
'  ③MailTo         : 宛先アドレス(複数の場合はカンマで区切る)
'  ④MailCc         : CCアドレス(複数の場合はカンマで区切る)
'  ⑤MailBcc        : BCCアドレス(複数の場合はカンマで区切る)
'  ⑥MailSubject    : 件名
'  ⑦MailBody       : 本文(改行はvbCrLf付加)
'  ⑧MailAddFile    : 添付ファイル(複数の場合はカンマで区切るか配列渡し) ※Option
'  ⑨MailCharacter  : 文字コード指定(デフォルトはShift-JIS)              ※Option
' [戻り値]
'  正常時："OK", エラー時："NG"+エラーメッセージ
'*******************************************************************************
Public Function SendMailCDO(strDialUp As String, MailSmtpServer As String, MailFrom As String, MailTo As String, MailCc As String, MailBcc As String, MailSubject As String, MailBody As String, Optional MailAddFile As Variant, Optional MailCharacter As String)
    Const cnsOK = "OK"
    Const cnsNG = "NG"
    Dim objCDO As New CDO.Message
    Dim vntFILE As Variant
    Dim IX As Long
    Dim strCharacter As String, strBody As String, strChar As String
    
'    On Error GoTo SendMailByCDO_ERR
    SendMailCDO = cnsNG
    
    Dim strDIAL_ENTRY As String     ' ﾀﾞｲｱﾙｱｯﾌﾟｴﾝﾄﾘｰ
    Dim strMSG As String            ' ﾒｯｾｰｼﾞ
    Dim swLine As Byte              ' ﾀﾞｲｱﾙ接続ｽｲｯﾁ
    Dim hWnd As Long                ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ
    Dim lngConnID As Long           ' ｺﾈｸｼｮﾝｺｰﾄﾞ
    Dim lngRet As Long              ' ﾘﾀｰﾝｺｰﾄﾞ
    
'    ' ﾀﾞｲｱﾙｱｯﾌﾟ接続(ｴﾝﾄﾘ名が指定されている場合のみ)
'    strDIAL_ENTRY = Trim$(strDialUp)
'    If strDIAL_ENTRY <> "" Then
'        strMSG = ""
'        swLine = 0
'        MainForm.StatusBar1.Panels(1) = "｢" & strDIAL_ENTRY & "｣に接続中です．．．．"
'        ' ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙを取得
'        hWnd = MainForm.hWnd 'FP_GET_HWND(strCaption)
'        lngConnID = 0
'        ' ﾘﾓｰﾄ接続を起動
'        lngRet = InternetDial(hWnd, strDIAL_ENTRY, INTERNET_AUTODIAL_FORCE_UNATTENDED, lngConnID, 0&)
'        If ((lngRet <> 0) And (lngRet <> 633)) Then
'            strMSG = "｢" & strDIAL_ENTRY & "｣への接続に失敗しました。"
'            Select Case lngRet
'                Case 623: strMSG = strMSG & vbCr & "　(ﾀﾞｲｱﾙｴﾝﾄﾘｰ名が不存在)"
'                Case 668: strMSG = strMSG & vbCr & "　(ﾊﾟｽﾜｰﾄﾞが未登録)"
'                Case Else: strMSG = strMSG & vbCr & "　(その他ｴﾗｰ : " & CStr(lngRet) & " )"
'            End Select
'            SendMailCDO = strMSG
'            Exit Function
'        End If
'        swLine = 1
'    Else
'        Exit Function
'    End If
    
    ' 文字コード指定の確認
    If MailCharacter <> "" Then
        ' 指定ありの場合は指定値をセット
        strCharacter = MailCharacter
    Else
        ' 指定なしの場合はShift-JISとする
        strCharacter = cdoShift_JIS
    End If
    
    ' 本文の改行コードの確認
    ' Lfのみの場合Cr+Lfに変換
    strBody = Replace(MailBody, vbLf, vbCrLf)
    ' 上記で元がCr+Lfの場合Cr+Cr+LfになるのでCr+Lfに戻す
    MailBody = Replace(strBody, vbCr & vbCrLf, vbCrLf)
    
    With objCDO
'        With .Configuration.Fields                          ' 設定項目
'            .Item(cdoSendUsingMethod) = cdoSendUsingPort    ' 外部SMTP指定
'            .Item(cdoSMTPServer) = MailSmtpServer           ' SMTPサーバ名
'            .Item(cdoSMTPServerPort) = 25                   ' ポート№
'            .Item(cdoSMTPConnectionTimeout) = 60            ' タイムアウト
'            .Item(cdoSMTPAuthenticate) = cdoAnonymous       ' 0
'            .Item(cdoLanguageCode) = strCharacter           ' 文字セット指定
'            .Update                                         ' 設定を更新
'        End With
        'SMTP認証ならこっち
'        strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
        With .Configuration.Fields
            .Item(cdoSendUsingMethod) = 2               ' 外部SMTP指定
            .Item(cdoSMTPServer) = MailSmtpServer      ' SMTPサーバ名
            .Item(cdoSMTPServerPort) = 465            ' ポート№
            .Item(cdoSMTPUseSSL) = True               ' SSLを使う場合にTrue
            .Item(cdoSMTPAuthenticate) = 1            ' 1(Basic認証)/2（NTLM認証）
            .Item(cdoSendUserName) = "atic.alertmail@gmail.com"
            .Item(cdoSendPassword) = "idappe99"
            .Item(cdoSMTPConnectionTimeout) = 60          ' タイムアウト
            .Item(cdoLanguageCode) = strCharacter           ' 文字セット指定
            .Update
        End With
        .Fields("urn:schemas:mailheader:X-Mailer") = "CDO mail"
        .Fields("urn:schemas:mailheader:Importance") = "High"
        .Fields("urn:schemas:mailheader:Priority") = 1
        .Fields("urn:schemas:mailheader:X-Priority") = 1
        .Fields("urn:schemas:mailheader:X-MsMail-Priority") = "High"
        .Fields.Update
        
        .MimeFormatted = True
        .Fields.Update
        .From = MailFrom                        ' 送信者
        .To = MailTo                            ' 宛先
        If MailCc <> "" Then .CC = MailCc       ' CC
        If MailBcc <> "" Then .BCC = MailBcc    ' BCC
        .Subject = MailSubject                  ' 件名
        .TextBody = MailBody                    ' 本文
        .TextBodyPart.Charset = strCharacter    ' 文字セット指定(本文)
        ' 添付ファイルの登録(複数対応)
        If ((VarType(MailAddFile) <> vbError) And (VarType(MailAddFile) <> vbBoolean) And (VarType(MailAddFile) <> vbEmpty) And (VarType(MailAddFile) <> vbNull)) Then
            If IsArray(MailAddFile) Then
                For IX = LBound(MailAddFile) To UBound(MailAddFile)
                    .AddAttachment MailAddFile(IX)
                Next IX
            ElseIf MailAddFile <> "" Then
                vntFILE = Split(CStr(MailAddFile), ",")
                For IX = LBound(vntFILE) To UBound(vntFILE)
                    If Trim(vntFILE(IX)) <> "" Then
                        .AddAttachment Trim(vntFILE(IX))
                    End If
                Next IX
            End If
        End If
        .Send                                   ' 送信
    End With
    Set objCDO = Nothing
    SendMailCDO = cnsOK

    ' ﾀﾞｲｱﾙｺﾈｸｼｮﾝを切断
    If swLine = 1 Then
        ' 回線を切断
        MainForm.StatusBar1.Panels(1) = "｢" & strDIAL_ENTRY & "｣を切断中です．．．．"
        InternetHangUp lngConnID, 0&
        swLine = 0
    End If
Dim strRC
    If strRC <> "" Then
        SendMailCDO = strRC & vbCr & vbCr & "サーバーに接続できないか、切断されました。"
        MainForm.StatusBar1.Panels(1) = False
        Exit Function
    End If




Exit Function

'-------------------------------------------------------------------------------
SendMailByCDO_ERR:
    SendMailCDO = cnsNG & Err.Number & " " & Err.Description
    On Error Resume Next
    Set objCDO = Nothing
End Function
