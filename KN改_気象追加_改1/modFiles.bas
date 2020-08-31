Attribute VB_Name = "modFiles"
Option Explicit

'ファイルハンドルを取得する
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

'ファイルから読み込む
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

'ファイルハンドルを閉じる
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const GENERIC_ALL = &H10000000
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const OPEN_EXISTING = 3

'ファイルタイム
Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
     dwFileAttributes As Long       'ファイル属性
     ftCreationTime As FILETIME     '作成日
     ftLastAccessTime As FILETIME   'アクセス日
     ftLastWriteTime As FILETIME    '更新日
     nFileSizeHigh As Long          'ファイルサイズ(Byte)
     nFileSizeLow As Long           'ファイルサイズ(Byte)
     dwReserved0 As Long            '未使用
     dwReserved1 As Long            '未使用
     cFileName As String * 260      'ファイル名
     cAlternate As String * 14      'ファイル名フォーマット名
End Type
'ファイルの検索を開始する
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
'ファイルの検索を続行する
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'検索ハンドルを閉じる
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Type DIR_FILE_LIST
    FILENAME As String
    IsDirectory As Boolean
End Type

'パス操作用
Private Declare Function PathFindFileName Lib "SHLWAPI.DLL" Alias "PathFindFileNameA" _
                                (ByVal pszPath As String) As Long
Private Const MAX_PATH = 260
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                (pDest As Any, _
                                 pSource As Any, _
                                 ByVal ByteLen As Long)

Private Declare Function PathRemoveBackslash Lib "SHLWAPI.DLL" Alias "PathRemoveBackslashA" _
                                (ByVal pszPath As String) As Long

Private Declare Function PathRemoveFileSpec Lib "SHLWAPI.DLL" Alias "PathRemoveFileSpecA" _
                                (ByVal pszPath As String) As Long


' ファイル操作用のAPI-----------------------------------------------------
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_DELETE = &H3&              ' 削除
Private Const FO_COPY = &H2                 ' コピー
Private Const FO_MOVE = &H1                 ' 移動
Private Const FO_RENAME = &H4               ' ファイル名変更
Private Const FOF_ALLOWUNDO = &H40&         ' ごみ箱へ
Private Const FOF_NOCONFIRMATION = &H10&    ' 確認ダイアログを表示しなし
Private Const FOF_NOERRORUI = &H400&        ' エラーダイアログを表示しない
Private Const FOF_MULTIDESTFILES = &H1&     ' 複数ファイルを指定する

'Private Const FO_MOVE As Long = &H1
'Private Const FO_COPY As Long = &H2
'Private Const FO_DELETE As Long = &H3
'Private Const FO_RENAME As Long = &H4
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
'Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
'Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
'Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200
'Private Const FOF_NOERRORUI As Long = &H400

' ファイルを検索する。
' RootPath          : 検索を開始する基準のディレクトリ
' InputPathName     : 検索するファイル名
' OutputPathBuffer  : 見つかったファイル名を格納するバッファ。
' 戻り値            : 見つかると0以外を返す。
Private Declare Function SearchTreeForFile Lib "imagehlp.dll" _
    (ByVal RootPath As String, _
     ByVal InputPathName As String, _
     ByVal OutputPathBuffer As String) As Long

Private Const MAX_PATHp = 512
Private Const MAX_PATH_PLUS1 = MAX_PATHp + 1

Dim ttt As String

'Sub main()
'FindDataFile "E:\A-TiC\G_業務\K_計測部\y_四ツ峰\ネットワーク対応\UPLOAD\TEST"
'End Sub

Public Function CheckDataFile0(ByVal fdir As String) As Long
'ファイル検索
'ファイルがあったら、配列にパス名を取得
'fDir : 検索ディレクトリ

    On Error GoTo CheckDataFile9999

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

    '    If aIndex = -1 Then Exit Function
    End If
    CheckDataFile0 = aIndex + 1
    Exit Function
    
CheckDataFile9999:
    CheckDataFile0 = 0
End Function

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

    If GetTargetFiles(FileList, fdir, "dat") Then
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
        
        'FTPで送信
        Dim rc As Integer
'        If id = 1 Then
''            Call SendFTP(rc, fdir, tFilename(), TdsDataPath) ' こっち間違い
'            Call SendFTP(rc, fdir, tFilename(), TDSFTPpath) '
'        End If
'        If id = 2 Then
'            Call SendFTP(rc, fdir, tFilename(), TDSFTPpath(id2))
'        End If
        
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


'-------------------------------------------------------------------
' 関数名 ： ReadFileUsingAPIFunc
' 機能 ： ファイルから文字列を読み込み テキスト型変数に代入する
' 引数 ： ×(in) srcText … 文字列を表示するテキストボックス
'           (in) fPath … 読み込むファイルのパス
' 返り値 ：  true : 読み取り成功
'           False : 読めなかった
'-------------------------------------------------------------------
Private Function ReadFileUsingAPIFunc(ByVal fPath As String) As Boolean

    Dim hFile As Long       'ファイルのハンドル
    Dim FileSize As Long    'ファイルサイズ
    Dim gBinData() As Byte  '取得データ
    Dim outFileSize As Long

    'ファイルを開く(READ)
    hFile = CreateFile(fPath, GENERIC_READ, 0&, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile = -1 Then
        ReadFileUsingAPIFunc = False
        Exit Function
    Else
    End If

    'ファイルサイズ取得
    FileSize = FileLen(fPath)
    If FileSize < 567 Then
        'ファイルを閉じる
        Call CloseHandle(hFile)
        ReadFileUsingAPIFunc = False
    End If

    '変数初期化
    ReDim Preserve gBinData(FileSize - 1) As Byte

    'ファイル読み込み
    Call ReadFile(hFile, gBinData(0), FileSize, outFileSize, 0&)

    'ANSI → Unicode変換
'    srcText.Text = StrConv(gBinData(), vbUnicode)
    ttt = StrConv(gBinData(), vbUnicode)

    'ファイルを閉じる
    Call CloseHandle(hFile)
    
    ReadFileUsingAPIFunc = True

End Function

'-----------------------------------------------------------------------
' 関数名 ： GetTargetFiles
' 機能   ： ディレクトリ以下の指定拡張子のファイルを取得する
' 引数   ： (in/out) Files … 取得したファイルを格納する配列
'           (in)DirName    … 検索元ディレクトリ
'           (in)Extension  … 拡張子
' 返り値 ： True：検索元ディレクトリは存在する  False：存在しない
'-----------------------------------------------------------------------
Public Function GetTargetFiles(ByRef Files() As String, _
                                ByVal DirName As String, _
                                ByVal Extension As String) As Boolean

    Dim udtWin32 As WIN32_FIND_DATA
    Dim hFile As Long
    Dim ArrayIndex As Long
    Dim FileListNum As Long
    Dim i As Long
    Dim UdtDFL() As DIR_FILE_LIST

    '最後尾に \ が付いている場合は取る
    If Right$(DirName, 1) = "\" Then DirName = Left$(DirName, Len(DirName) - 1)

    '検索開始
    hFile = FindFirstFile(DirName, udtWin32)
    'ファイル・ディレクトリが存在しない場合は終了
    If hFile = INVALID_HANDLE_VALUE Then
        WriteLog DirName & " - ファイル・ディレクトリが存在しない"
        Exit Function
    End If
    Call FindClose(hFile)

    'ディレクトリ以下のファイル・ディレクトリを取得する
    FileListNum = GetFileList(UdtDFL, DirName)
    If FileListNum = (-1) Then Exit Function

    For i = 0 To FileListNum
        'ディレクトリである
        If UdtDFL(i).IsDirectory Then
            Call GetTargetFiles(Files, DirName & "\" & UdtDFL(i).FILENAME, Extension)
        'ファイルである
        Else
            If UCase$(Left(UdtDFL(i).FILENAME, 1)) = "R" Then
                'ファイルの拡張子が指定拡張子と等しい
                If UCase$(Right$(UdtDFL(i).FILENAME, Len(Extension))) = UCase$(Extension) Then
                    '初回実行時 Files は配列無しなのでUBound()でエラーとなる
                    'それを回避するための強制エラー無視ロジック
                    On Error Resume Next
                    ArrayIndex = UBound(Files) + 1
                    On Error GoTo 0
    
                    'メモリー確保
                    ReDim Preserve Files(ArrayIndex) As String
    
                    'フルパスでファイル名を格納
                    Files(ArrayIndex) = DirName & "\" & UdtDFL(i).FILENAME
                End If
            End If
        End If
    Next i

    GetTargetFiles = True
    On Error GoTo 0

End Function

'-----------------------------------------------------------------------
' 関数名 ： GetFileList
' 機能   ： ディレクトリのファイルを取得する
' 引数   ： (in/out) UdtDFL … DIR_FILE_LIST構造体の配列
'           (in)DirName     … 検索元ディレクトリ
' 返り値 ： UdtDFL配列数   ファイルが存在しない場合：-1
'-----------------------------------------------------------------------
Private Function GetFileList(ByRef UdtDFL() As DIR_FILE_LIST, _
                            ByVal DirName As String) As Long

    Dim udtWin32 As WIN32_FIND_DATA
    Dim hFile As Long
    Dim ArrayIndex As Long
    Dim Win32FileName As String

    ArrayIndex = (-1)

    '検索開始
    hFile = FindFirstFile(DirName & "\*", udtWin32)
    Do
        '時々、再描画
        If ArrayIndex Mod 10 = 0 Then DoEvents

        'ファイル名取得
        Win32FileName = Left$(udtWin32.cFileName, _
                              InStr(udtWin32.cFileName, Chr$(0)) - 1)

        '親ディレクトリ、カレントディレクトリでない
        If Left$(Win32FileName, 1) <> "." Then
            ArrayIndex = ArrayIndex + 1
            ReDim Preserve UdtDFL(ArrayIndex) As DIR_FILE_LIST
            'ファイル名、ファイル属性を取得
            With UdtDFL(ArrayIndex)
                .FILENAME = Win32FileName
                .IsDirectory = CBool(udtWin32.dwFileAttributes And vbDirectory)
            End With
        End If
    Loop While FindNextFile(hFile, udtWin32) <> 0

    Call FindClose(hFile)

    GetFileList = ArrayIndex

End Function

Public Sub s_ShellSort(ByRef sArray() As String, ByVal Num As Integer)
   Dim Span As Integer
   Dim i As Integer
   Dim j As Integer
   Dim TMP As String
   
   Span = Num \ 2
   Do While Span > 0
      For i = Span To Num - 1
         j% = i% - Span + 1
         For j = (i - Span + 1) To 0 Step -Span
            If sArray(j) <= sArray(j + Span) Then Exit For
            ' 順番の異なる配列要素を入れ替えます.
            TMP = sArray(j)
            sArray(j) = sArray(j + Span)
            sArray(j + Span) = TMP
         Next j
      Next i
      Span = Span \ 2
   Loop
End Sub
'
' ファイル名を取り出す。
'
Public Function FindFileName(ByVal strFileName As String) As String
    ' strFileName : フルパスのファイル名
    ' 戻り値      : ファイル名だけが返る。
    Dim strBuffer   As String
    Dim lngResult   As Long
    Dim bytStr()    As Byte

    lngResult = PathFindFileName(strFileName)
    If lngResult <> 0 Then
        ' (MAX_PATH + 1)のバイト配列を用意する。
        ReDim bytStr(MAX_PATH + 1) As Byte
        ' 確保したバイト配列に得られた位置のメモリをコピーする。
        MoveMemory bytStr(0), ByVal lngResult, MAX_PATH + 1
        ' 配列を文字列に変換する。
        strBuffer = StrConv(bytStr(), vbUnicode)
        ' NULL文字までを切り出す。
        FindFileName = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    End If
End Function

'
' パスだけを取り出す。
'
Public Function RemoveFileSpec(ByVal strPath As String) As String
    ' strPath : フルパスのファイル名
    ' 戻り値  : パス名

    Dim lngResult As Long
    lngResult = PathRemoveFileSpec(strPath)
    If lngResult <> 0 Then
        If InStr(strPath, vbNullChar) > 0 Then
            RemoveFileSpec = Left$(strPath, InStr(strPath, vbNullChar) - 1)
        Else
            RemoveFileSpec = strPath
        End If
    End If
End Function

Public Sub sFileDelete(DelFile As String)
    '**************************************************************
    '* SHFileOperation関数を呼び出しファイルをごみ箱に送る　　　　*
    '* meForm　 = ダイアログを表示するForm　　　　　　　　　　　　    *
    '* DelFile　= 削除するファイル名（Path付）　　　　　　　　　　    *
    '*　　　　　　複数のファイルを指定する場合vbNullCharで区切り　  *
    '*　　　　　　最後は二つのvbNullCharで終わる　　　　　　　　　    *
    '**************************************************************
    On Error Resume Next
    Dim lpFileOp As SHFILEOPSTRUCT
    Dim Result   As Long
    Dim MyFlag   As Long

'ゴミ箱の場合
    '指定方法はお好みで設定して下さい。
    MyFlag = FOF_ALLOWUNDO                  'ごみ箱へ
    MyFlag = MyFlag + FOF_NOCONFIRMATION    '確認しない
    ''MyFlag = MyFlag + FOF_MULTIDESTFILES    '複数ファイル
    MyFlag = MyFlag + FOF_NOERRORUI         'エラーのダイアログを非表示

    ' ファイル操作に関する情報を指定
    With lpFileOp
        .hWnd = App.hInstance ' .hWnd       ' ダイアログの親ウィンドウハンドルを指定
        .wFunc = FO_DELETE       ' 削除を指定
        .pFrom = DelFile         ' 削除するディレクトリを指定
       ' .pTo = 操作先のファイル名・ディレクトリ名
        .fFlags = MyFlag         '動作方法を指定
    End With

    ' ファイル操作を実行
    Result = SHFileOperation(lpFileOp)
    On Error GoTo 0

End Sub

Public Sub sFileMove(DelFile As String)
    '**************************************************************
    '* SHFileOperation関数を呼び出しファイルをごみ箱に送る　　　　*
    '* meForm　 = ダイアログを表示するForm　　　　　　　　　　　　*
    '* DelFile　= 削除するファイル名（Path付）　　　　　　　　　　*
    '*　　　　　　複数のファイルを指定する場合vbNullCharで区切り　*
    '*　　　　　　最後は二つのvbNullCharで終わる　　　　　　　　　*
    '**************************************************************
    On Error Resume Next
    Dim lpFileOp As SHFILEOPSTRUCT
    Dim Result   As Long
    Dim MyFlag   As Long

'ゴミ箱の場合
'    '指定方法はお好みで設定して下さい。
'    MyFlag = FOF_ALLOWUNDO                  'ごみ箱へ
'    MyFlag = MyFlag + FOF_NOCONFIRMATION    '確認しない
'    ''MyFlag = MyFlag + FOF_MULTIDESTFILES  '複数ファイル
'    MyFlag = MyFlag + FOF_NOERRORUI         'エラーのダイアログを非表示
'
'    ' ファイル操作に関する情報を指定
'    With lpFileOp
'        .hWnd = App.hInstance ' .hWnd       ' ダイアログの親ウィンドウハンドルを指定
'        .wFunc = FO_DELETE       ' 削除を指定
'        .pFrom = DelFile         ' 削除するディレクトリを指定
'       ' .pTo = 操作先のファイル名・ディレクトリ名
'        .fFlags = MyFlag         '動作方法を指定
'    End With

    MyFlag = FOF_NOCONFIRMMKDIR                  'ごみ箱へ
    MyFlag = MyFlag + FOF_NOCONFIRMATION    '確認しない
    ''MyFlag = MyFlag + FOF_MULTIDESTFILES    '複数ファイル
    MyFlag = MyFlag + FOF_NOERRORUI         'エラーのダイアログを非表示

    ' ファイル操作に関する情報を指定
    With lpFileOp
        .hWnd = App.hInstance ' .hWnd       ' ダイアログの親ウィンドウハンドルを指定
        .wFunc = FO_MOVE       ' 削除を指定
        .pFrom = DelFile         ' 削除するディレクトリを指定
        .pTo = App.Path & "\tmp\" '操作先のファイル名・ディレクトリ名
        .fFlags = MyFlag         '動作方法を指定
    End With

    ' ファイル操作を実行
    Result = SHFileOperation(lpFileOp)
    
    On Error GoTo 0
    
End Sub

'Public Sub SendFTP(ret As Integer, fdir As String, fPath() As String, FTPpath$)
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
'    On Local Error GoTo SendFTPerr
'
'    Dim ftpErr  As String
'    Dim rc As Long
'    Dim vv As Variant, vv2 As Variant
'''    Dim ftp As Object
'''    Set ftp = CreateObject("basp21.FTP")
'    Dim ftp As BASP21Lib.ftp
'    Set ftp = New BASP21Lib.ftp
'
'    ftp.OpenLog App.Path & "\FTP-log.txt"
''    rc = ftp.Connect("172.16.60.219", "a-tic", "keisoku")  '本物
'If Command$ = "TEST" Then
'    rc = ftp.Connect(ServerIP, "anonymous", "")  '
'Else
''    ServerIP = "153.150.115.38"
'    ServerIP = Atesaki '"180.43.16.132"
'    rc = ftp.Connect(ServerIP, sUser, "ftp_keisoku")  '本物
'End If
''    rc = ftp.Connect("60.43.239.36", "chikah", "zbeba+nn")  '前の本物
'    If rc = 0 Then
'        '計測データのアップロード
'        For i = 0 To UBound(fPath)
'            tFile = FTPpathname(fPath(i), sYY, sMM, sDD)
'            If Left$(tFile, 1) = "/" Then tFile = Right$(tFile, Len(tFile) - 1)
'           '格納先ディレクトリが無かったら作成してから移動する
'                        rc = ftp.Command("CWD " & FTPpath & "/" & sYY)   'ディレクトリ移動
'            If Not (rc = 2) Then
'               '移動に失敗した場合、年のディレクトリがなかったと仮定する
'                                ftpErr = ftp.GetReply()
'               'If InStr(ftpErr, "No such file or directory") > 0 Then
'                If InStr(ftpErr, "directory not found") > 0 Or _
'                   InStr(ftpErr, "No such file or directory") > 0 Or _
'                   InStr(ftpErr, "Failed to change directory.") > 0 Then
'                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY)                            'ディレクトリ作成
'                    If Not (rc = 2) Then
'                        'MKDIRができない時は、以降の処理が不能なので中断する。
'                            Exit For
'                    End If
'                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)                'ディレクトリ作成
'                    rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
'                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
'                End If
'            Else
'                rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM)    'ディレクトリ移動
'                If Not (rc = 2) Then
'                   '月のディレクトリがなかった場合
'                    ftpErr = ftp.GetReply()
'                    'If InStr(ftpErr, "No such file or directory") > 0 Then
'                    If InStr(ftpErr, "directory not found") > 0 Or _
'                       InStr(ftpErr, "No such file or directory") > 0 Or _
'                       InStr(ftpErr, "Failed to change directory.") > 0 Then
'                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM)                'ディレクトリ作成
'                        If Not (rc = 2) Then
'                                'MKDIRができない時は、以降の処理が不能なので中断する。
'                                Exit For
'                        End If
'                        rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
'                        rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
'                    End If
'                Else
'                    rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)        'ディレクトリ移動
'                    If Not (rc = 2) Then
'                       '日のディレクトリがなかった場合
'                        ftpErr = ftp.GetReply()
'                        'If InStr(ftpErr, "No such file or directory") > 0 Then
'                        If InStr(ftpErr, "directory not found") > 0 Or _
'                           InStr(ftpErr, "No such file or directory") > 0 Or _
'                           InStr(ftpErr, "Failed to change directory.") > 0 Then
'                            rc = ftp.Command("MKD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ作成
'                            If Not (rc = 2) Then
'                                    'MKDIRができない時は、以降の処理が不能なので中断する。
'                                    Exit For
'                            End If
'                            rc = ftp.Command("CWD " & FTPpath & "/" & sYY & "/" & sMM & "/" & sDD)    'ディレクトリ移動
'                        End If
'                    End If
'                End If
'            End If
'           'カレントディレクトリが移動済、そこにファイル送信
'            rc = ftp.PutFile(fdir & "\" & fPath(i), "", 1) 'ファイル送信
'
'           '送信されたか確認
'            If rc = 1 Then
'                fpSW = False
'                vv = ftp.GetDir("") ' ディレクトリ一覧(ファイル名)
'                If IsArray(vv) Then
'                    For Each vv2 In vv
'                        If vv2 = fPath(i) Then
'                           '格納先に送信したファイルがあった!!
'                                                        fpSW = True
'                            Exit For
'                        End If
'                    Next
'                End If
'                If fpSW = True Then
'                    '送信成功の場合、元のファイルを削除(ごみ箱行き)
'                    sFileDelete fdir & "\" & fPath(i)
'                End If
'            End If
'        Next i
'        ftp.Close
'    Else
'        ftpErr = ftp.GetReply()
'    End If
'    rc = ftp.CloseLog()
'
'    Set ftp = Nothing
'    ret = -1
'Exit Sub
'SendFTPerr:
'    Set ftp = Nothing
'End Sub
'
