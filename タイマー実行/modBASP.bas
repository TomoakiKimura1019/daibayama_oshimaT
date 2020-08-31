Attribute VB_Name = "modBASP"
Option Explicit

Public Sub FindDataFile(ByVal id%, ByVal fdir As String, ByVal id2 As Integer)
'ファイル検索
'ファイルがあったら、配列にパス名を取得
'fDir : 検索ディレクトリ

    Dim FileList() As String
    Dim i As Long

    Dim ret As String

    Dim tFilename() As String
    Dim aIndex As Long
    aIndex = -1

        If GetTargetFiles(FileList, fdir, "csv") Then
        'ファイル名を配列に取得
        For i = 0 To UBound(FileList)
'            Debug.Print FileList(i)
            '所定の型式のファイルを選択
            'ret = Match("/\d{1,4}_\d{1,2}_BV\d{1}-[XY]_disp.txt/", FindFileName(FileList(i)))
'            ret = Match("/\d{1,4}_\d{1,2}_strain.txt/", FindFileName(FileList(i)))
'            If ret = "1" Then
                aIndex = aIndex + 1
                ReDim Preserve tFilename(aIndex) As String
                tFilename(aIndex) = FindFileName(FileList(i))
'            End If
        Next i
        '所得したファイル名をソート
        If -1 < aIndex Then
            s_ShellSort tFilename(), (aIndex)
        End If

        If aIndex = -1 Then Exit Sub
        
        'frmTDSdataget.StatusBar1.Panels(1).Text = "found"
        
        'FTPで送信
        Dim rc As Integer
'        If id = 1 Then
''            Call SendFTP(rc, fdir, tFilename(), TdsDataPath) ' こっち間違い
'            Call SendFTP(rc, fdir, tFilename(), TDSFTPpath) '
'        End If
        If id = 2 Then
            'frmTDSdataget.StatusBar1.Panels(1).Text = "SendFTP start"
            Call SendFTP(rc, fdir, tFilename(), "/array1/share2/共有書庫/計測部員の書庫/etc")
        End If
        
    End If

End Sub

Public Function FTPpathname(ByVal tFilename As String, sYY$, sMM$, sDD$) As String
'ファイル名から目的のFTPディレクトリ名を生成

'    Dim sYY As String
'    Dim sMM As String
'    Dim sDD As String
    Dim sNN As String
    
    '2009-10-12_10-00.dat
    sYY = Mid$(tFilename, 1, 4)
    sMM = Mid$(tFilename, 6, 2)
    sDD = Mid$(tFilename, 9, 2)
    sNN = "/" & sYY & "/" & sMM & "/" & sDD
    
    FTPpathname = sNN
    
End Function

Public Sub SendFTP(ret As Integer, fdir As String, fPath() As String, FTPpath$)
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
    
    Dim ServerIP As String
    Dim i As Integer
    Dim tFile As String
    
    Dim sYY As String
    Dim sMM As String
    Dim sDD As String
    Dim fpSW As Boolean
    ret = 0
    On Local Error GoTo SendFTPerr
    
    Dim ftpErr  As String
    Dim rc As Long
    Dim vv As Variant, vv2 As Variant
''    Dim ftp As Object
''    Set ftp = CreateObject("basp21.FTP")
    Dim ftp As BASP21Lib.ftp
    Set ftp = New BASP21Lib.ftp
    
    ftp.OpenLog App.Path & "\FTP-log.txt"
'    rc = ftp.Connect("172.16.60.219", "a-tic", "keisoku")  '本物
If Command$ = "TEST" Then
    ServerIP = "172.16.60.99"
    rc = ftp.Connect(ServerIP, "anonymous", "")  '
Else
    ServerIP = "180.43.16.132"
'    ServerIP = "172.16.65.96"
    rc = ftp.Connect(ServerIP & ":49621", "otonaka", "atic2803")  '本物
End If
'    rc = ftp.Connect("60.43.239.36", "chikah", "zbeba+nn")  '前の本物
    If rc = 0 Then
        'frmTDSdataget.StatusBar1.Panels(1).Text = "FTP connect"
        ' passiveモードにする
        ftp.Command ("PASV") ' 一度呼出せば OK
        '計測データのアップロード
        For i = 0 To UBound(fPath)
            tFile = FTPpathname(fPath(i), sYY, sMM, sDD)
            If Left$(tFile, 1) = "/" Then tFile = Right$(tFile, Len(tFile) - 1)
            rc = ftp.Command("CWD " & FTPpath & "/" & sYY)   'ディレクトリ移動
            If Not (rc = 2) Then
                ftpErr = ftp.GetReply()
                'If InStr(ftpErr, "No such file or directory") > 0 Then
                If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY)    'ディレクトリ作成
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)    'ディレクトリ作成
                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
                End If
            Else
                rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM)    'ディレクトリ移動
                If Not (rc = 2) Then
                    ftpErr = ftp.GetReply()
                    'If InStr(ftpErr, "No such file or directory") > 0 Then
                    If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)    'ディレクトリ作成
                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
                        rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
                    End If
                Else
                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
                    If Not (rc = 2) Then
                        ftpErr = ftp.GetReply()
                        'If InStr(ftpErr, "No such file or directory") > 0 Then
                        If InStr(ftpErr, "directory not found") > 0 Or InStr(ftpErr, "No such file or directory") > 0 Then
                            rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
                            rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
                        End If
                    End If
                End If
            End If
            
            rc = ftp.PutFile(fdir & "\" & fPath(i), "", 1) 'ファイル送信
            
            If rc = 1 Then
                fpSW = False
                vv = ftp.GetDir("") ' ディレクトリ一覧(ファイル名)
                If IsArray(vv) Then
                    For Each vv2 In vv
                        If vv2 = fPath(i) Then
                            fpSW = True
                            Exit For
                        End If
                    Next
                End If
                If fpSW = True Then
                    sFileDelete fdir & "\" & fPath(i)
                End If
            End If
        Next i
        ftp.Close
        'frmTDSdataget.StatusBar1.Panels(1).Text = ""
    Else
        ftpErr = ftp.GetReply()
        'frmTDSdataget.StatusBar1.Panels(1).Text = "FTP connect error"
    End If
    rc = ftp.CloseLog()
    
    Set ftp = Nothing
    ret = -1
Exit Sub
SendFTPerr:
    Set ftp = Nothing
End Sub

