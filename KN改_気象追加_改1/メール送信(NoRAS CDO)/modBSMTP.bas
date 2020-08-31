Attribute VB_Name = "modBSMTP"
Option Explicit

'
' 参照設定でBSMTPにチェックを入れる
'
'------------------------------------------------------
Private Declare Function SendMail Lib "bsmtp" _
      (szServer As String, szTo As String, szFrom As String, _
      szSubject As String, szBody As String, szFile As String) As String
Private Declare Function RcvMail Lib "bsmtp" _
      (szServer As String, szUser As String, szPass As String, _
      szCommand As String, szDir As String) As Variant
Private Declare Function ReadMail Lib "bsmtp" _
      (szFilename As String, szPara As String, szDir As String) As Variant

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

'FTPサーバ用
Public Type FTPsv
    Name As String
    User As String
    Pass As String
End Type

Public mINIfile As String
'Public strData$

'
'##############################################################
Public Sub FTPdataGet(SV As FTPsv, ret As Integer)
'FTPサーバからデータダウンロード
'    On Local Error GoTo SendFTPerr
    
    Dim fileN() As String, fileC As Long, fileNS As String
    Dim ftpErr  As String

    Dim ftp As BASP21Lib.ftp
    Dim rc As Long
    Dim vv As Variant, vv2 As Variant
    
    Dim i&
    
'    Dim FTPdir$
    
    Set ftp = New BASP21Lib.ftp
    ftp.OpenLog App.Path & "\FTP-log.txt"
    rc = ftp.Connect(SV.Name, SV.User, SV.Pass)
    If rc = 0 Then
        vv = ftp.GetDir("/A0002/appdata") ' ディレクトリ一覧(ファイル名のみ)
        If IsArray(vv) Then
            fileC = 0
            For Each vv2 In vv
                fileNS = vv2
                fileNS = Trim$(fileNS)
                If UCase$(Right$(fileNS, 4)) = ".DAT" Then
                    fileC = fileC + 1
                    ReDim Preserve fileN(fileC)
                    fileN(fileC) = vv2
                    Debug.Print (fileC), fileN(fileC)
                End If
            Next
        End If
        '計測データのダウンロード
        For i = 1 To fileC
            rc = ftp.GetFile("/A0002/appdata/" & fileN(i), App.Path & "\AppData")  ' テキストファイルの受信
            If rc = 1 Then
                rc = ftp.DeleteFile("/A0002/appdata/" & fileN(i))
            End If
        Next i
        
        ftp.Close
    Else
        ftpErr = ftp.GetReply()
    End If
    
    ret = fileC
FTPrw001:
    rc = ftp.CloseLog()
    Set ftp = Nothing
    Erase fileN
Exit Sub

SendFTPerr:
    Set ftp = Nothing
'    Call ErrLog(Now, "FTPrw", (Err.Number & " " & Err.Description))
    Erase fileN
End Sub

Public Sub FTPdataPUT(SV As FTPsv, ret As Integer)
'FTPサーバへデータアップロード
    On Local Error GoTo SendFTPerr
    
    Dim ftpErr  As String

    Dim ftp As BASP21Lib.ftp
    Dim rc As Long
    
    ret = 0
'    Dim FTPdir$
    
    Set ftp = New BASP21Lib.ftp
    ftp.OpenLog App.Path & "\FTP-log.txt"
    rc = ftp.Connect(SV.Name, SV.User, SV.Pass)
    If rc = 0 Then
        If FileExists(App.Path & "\FTP\master.dat") = True Then
            rc = ftp.PutFile(App.Path & "\FTP\master.dat", "/A0002/DATA", 2) 'ファイル送信
            If 1 = rc Then
                Call DelFile(App.Path & "\FTP\" & "MASTER.DAT")
                ret = -1
            End If
        End If
        
        ftp.Close
    Else
        ftpErr = ftp.GetReply()
    End If
    
FTPrw001:
    rc = ftp.CloseLog()
    Set ftp = Nothing
Exit Sub
        
SendFTPerr:
    Set ftp = Nothing
End Sub

Public Sub MailSend(MailTbl As MailType)
   
    Dim i As Integer
    With MailTbl
    
        .SendCO = CInt(GetIni("メール送信", "送信数", mINIfile))
        For i = 1 To .SendCO
            .SendName(i) = GetIni("メール送信", "送信先" & CStr(i), mINIfile)
        Next i
    End With
    
    Dim ssb As String
    ssb = (GetIni("メール送信", "subject", mINIfile))
    
    Dim ret As String
    Dim szServer As String, szTo As String, szFrom As String
    Dim szSubject As String, szBody As String, szFile As String
    
    szServer = MailTbl.ServerName '& ":465:60"    ' SMTPサーバ名。ポート番号を指定できます。
    szServer = "smtp.gmail.com:465:60"               ' SMTPサーバ名。ポート番号を指定できます。"
    szTo = "who@who.com"            ' 宛先 ' 複数の宛先に送付するときは、アドレスをタブで区切っていくらでも指定できます。
    szTo = MailTbl.SendName(1)
    If MailTbl.SendCO > 1 Then
        szTo = szTo & vbTab & "bcc"
    End If
    For i = 2 To MailTbl.SendCO
        szTo = szTo & vbTab & MailTbl.SendName(i)
    Next i
            '    ' CCを指定するには次のようにします。
            '        szTo = "who@who.com" & vbTab & "cc" & vbTab & "who2@who2.com" & _
            '           vbTab & "who3@who3.com"
            '    ' BCCを指定するには次のようにします。
            '        szTo = "who@who.com" & vbTab & "bcc" & vbTab & "who2@who2.com" & _
            '           vbTab & "who3@who3.com"
            '    ' ヘッダを指定するには次のようにタブで区切り、>をヘッダの前に
            '    ' つけます。
            '        szTo = "who@who.com" & vbTab & ">Message-ID: 12345"
'    szFrom = MailTbl.ClientMailAddress & vbTab & MailTbl.Clientname & ":" & MailTbl.mailPassword   ' 送信元
    szFrom = MailTbl.ClientMailAddress & vbTab & "a545352322" & ":" & "yo2803ks"   ' 送信元
    szFrom = "<mbkeisya@gmail.com>" & vbTab & "mbkeisya@gmail.com" & ":" & "keisoku2803"   ' 送信元
    szSubject = ssb '     ' 件名
    'szBody = "こんにちは。" & vbCrLf & "さようなら"   ' 本文' 本文内で改行するには、vbCrLfを使います。
    szBody = strData '"測定データ"
    
    ' ファイルを添付するときは、ファイル名をフルパスで指定します。
    ' ファイルを複数指定するときは、タブで区切ってください。
    
''    szFile = CurrentDIR & "WaveData.LZH" '& vbTab & "c:\a2.jpeg" ' ファイル２個
    ' ファイルを添付しないときは次のようにします。
    szFile = ""   ' ファイル添付なし
    
    ret = SendMail(szServer, szTo, szFrom, szSubject, szBody, szFile)
    
    ' 送信エラーのときは、戻り値にエラーメッセージが返ります。
    If Len(ret) <> 0 Then
       'MsgBox "エラー" & ret
       Call ErrLog(Now, "メール送信", ret)
    End If
Exit Sub

End Sub

Public Sub MailRead(MailTbl As MailType)
    Dim szServer As String, szUser As String, szPass As String
    Dim szCommand As String, szDir As String
    Dim ar As Variant, v As Variant

    Dim szFilename As String, szPara As String
    Dim retv As Variant
    
    Dim t1 As Date, t2 As Date
    t1 = Now
    Do
        t2 = Now
        If DateDiff("s", DateAdd("s", 2, t1), t2) > 0 Then Exit Do
    Loop
    szServer = MailTbl.ServerName  'SMTPサーバ名と同じでよい。
                                    'タブで区切ってポート番号を指定できます。
    szUser = MailTbl.Clientname    'メールアカウント名
    szPass = MailTbl.mailPassword  'パスワード
    '''      2000/05/20 APOPをサポート
    '''      APOP 認証をするには、パスワードの前に "a" または "A" に １個の
    '''      ブランクをつけます｡
    '''      "a xxxx" : サーバがAPOP 未対応なら通常のUSER/PASS 処理をします。
    '''      "A xxxx" : サーバがAPOP 未対応ならエラーになります。
       
#If DebugVersion Then
    szCommand = "SAVEALL"  'コマンド　メールの１件目から３件目までを受信
#Else
    szCommand = "SAVEALLD"  'コマンド　メールの１件目から３件目までを受信
#End If
    
    szDir = App.Path & "\MailData" 'MailTabl.savefolder '受信したメールを保存するディレクトリ
    
    ar = RcvMail(szServer, szUser, szPass, szCommand, szDir)
    
'    Dim smCMD$
    '戻り値が返る変数は、Variantタイプを指定すること。
    '受信したメール１通ごとにファイルが作成されます。
    'メールに添付されたファイルは、本文と共に１つのファイルに含まれます。
    'ReadMail関数で添付ファイルを取出します。
    If IsArray(ar) Then   '正常終了時のSAVEコマンドの戻り値は、配列になります。
        For Each v In ar
            'Debug.Print v     'メールデータが保存されたファイル名がフルパスで戻ります。
                              'このファイル名をReadMailのパラメータとして渡します。
            szFilename = v  ' ファイル名にはRcvMailの戻り値の配列からファイル名を設定
            szPara = "subject:from:date:"  ' ヘッダーの指定
                                           ' nofile: とすると添付ファイルを保存しません。
            retv = ReadMail(szFilename, szPara, szDir)
            If IsArray(retv) Then
'                If 0 < InStr(retv(0), "インターバル変更") Then
'                    smCMD = retv(3)
'                    Call gCMD(smCMD)
'                ElseIf 0 < InStr(retv(0), "警報発生") Then
'                    smCMD = retv(3)
'                    Call KeihouARM(smCMD)
'                End If
                
                'For Each v2 In retv
                '     Debug.Print v2
                'Next
            Else
                'Debug.Print retv
            End If
        Next
    Else
        'Debug.Print ar      'エラー発生時は、配列でなくメッセージが戻ります。
    End If

End Sub

