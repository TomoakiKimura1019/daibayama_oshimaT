Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExW" ( _
    ByVal hWnd As OLE_HANDLE, ByVal pszPath As Long, ByVal psa As Any) As Long

Private Const RAD As Double = 3.14159265358979 / 180#

Dim oPath As String   'TSのデータ格納Path
Dim oFile(3) As String   '自分で管理するのデータファイル名
Dim oFileA As String  'TSのデータファイル名に付く、日時以外の文字列
Dim tFile As String
Dim heFile As String   'TSのデータから変位に
Dim fUPDATE As Boolean

Dim LastFilename As String

Dim LastDate As String

Type zahyo
   id As Integer
    x As Double
    y As Double
    z As Double
End Type

Dim kakudo As Double  '座標回転角度 DEG

Dim Pname(40) As String   '測点名称
Dim PnameK(40) As String   '測点名称 警報用
Dim INIT(40) As zahyo     '初期座標 元
Dim dINIT(40) As zahyo    '初期座標 回転後
Dim offSET(40) As zahyo   '変位の補正量 mm


Dim sokutenName(4) As String
Dim kanriLV As Integer
Dim kanriV(3) As Double
Public LOGFILE As String

Public ALERTfile As String

Public KN_Path(2) As String
Public KN_table(3) As String
Public KN_Offset(3) As String
Public KN_PathBK(2) As String
Public sokutenSu(3) As Integer
Public KN_SubName(3) As String

Public SoushinPath(3) As String
Public SoushinPathZ(3) As String

Public SoushinWM(3) As String

Public GroupName(3) As String

Sub Main()
    
'test

    
    Dim i As Integer
    Dim j As Integer
    
    KN_Path(1) = GetIni("system", "KN_Path1", App.Path & "\TSdata.ini")
    KN_Path(2) = GetIni("system", "KN_Path2", App.Path & "\TSdata.ini")
    If Right$(KN_Path(1), 1) <> "\" Then KN_Path(1) = KN_Path(1) & "\"
    If Right$(KN_Path(2), 1) <> "\" Then KN_Path(2) = KN_Path(2) & "\"
    
    KN_PathBK(1) = GetIni("system", "KN_MovePath1", App.Path & "\TSdata.ini")
    KN_PathBK(2) = GetIni("system", "KN_MovePath2", App.Path & "\TSdata.ini")
    If Right$(KN_PathBK(1), 1) <> "\" Then KN_PathBK(1) = KN_PathBK(1) & "\"
    If Right$(KN_PathBK(2), 1) <> "\" Then KN_PathBK(2) = KN_PathBK(2) & "\"
    
    KN_table(1) = GetIni("system", "KN_table1", App.Path & "\TSdata.ini")
    KN_table(2) = GetIni("system", "KN_table2", App.Path & "\TSdata.ini")
    
    KN_Offset(1) = GetIni("system", "KN_Offset1", App.Path & "\TSdata.ini")
    KN_Offset(2) = GetIni("system", "KN_Offset2", App.Path & "\TSdata.ini")
    
    SoushinPath(1) = GetIni("system", "SendPath1", App.Path & "\TSdata.ini")
    SoushinPath(2) = GetIni("system", "SendPath2", App.Path & "\TSdata.ini")
    If Right$(SoushinPath(1), 1) <> "\" Then SoushinPath(1) = SoushinPath(1) & "\"
    If Right$(SoushinPath(2), 1) <> "\" Then SoushinPath(2) = SoushinPath(2) & "\"
    
    SoushinPathZ(1) = GetIni("system", "SendPath1z", App.Path & "\TSdata.ini")
    SoushinPathZ(2) = GetIni("system", "SendPath2z", App.Path & "\TSdata.ini")
    If Right$(SoushinPathZ(1), 1) <> "\" Then SoushinPathZ(1) = SoushinPathZ(1) & "\"
    If Right$(SoushinPathZ(2), 1) <> "\" Then SoushinPathZ(2) = SoushinPathZ(2) & "\"
    
    '気象用送信ディレクトリ
    SoushinWM(1) = GetIni("system", "sendWM1", App.Path & "\TSdata.ini")
    SoushinWM(2) = GetIni("system", "sendWM2", App.Path & "\TSdata.ini")
    If Right$(SoushinWM(1), 1) <> "\" Then SoushinWM(1) = SoushinWM(1) & "\"
    If Right$(SoushinWM(2), 1) <> "\" Then SoushinWM(2) = SoushinWM(2) & "\"
    
    'GroupName(1) = GetIni("Group", "Name1", App.Path & "\TSdata.ini")
    'GroupName(2) = GetIni("Group", "Name2", App.Path & "\TSdata.ini")
    
    oPath = GetIni("system", "oPath", App.Path & "\TSdata.ini")
    oFile(1) = GetIni("system", "oFile1", App.Path & "\TSdata.ini")
    oFile(2) = GetIni("system", "oFile2", App.Path & "\TSdata.ini")
    
    heFile = GetIni("system", "hFile", App.Path & "\TSdata.ini")
    oFileA = GetIni("system", "oFileA", App.Path & "\TSdata.ini")
    ALERTfile = GetIni("system", "ALERTfile", App.Path & "\TSdata.ini")
    
    sokutenSu(1) = GetIni("Group", "TenSu1", App.Path & "\TSdata.ini")
    sokutenSu(2) = GetIni("Group", "TenSu2", App.Path & "\TSdata.ini")
    
    KN_SubName(1) = "RAIL01"
    KN_SubName(2) = "RAIL02"
    
'            Call ALERTfileCK("2017/09/19 20:00:00", j)
    
    For i = 1 To 3
        kanriV(i) = GetIni("kanri", "Vkanri" & i, App.Path & "\TSdata.ini")
    Next i
    
    
    '"2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978
    
    LOGFILE = "TDdata.log"
    WriteLog "logfile"
    
    Dim st As String
    Dim id As Integer
    
    id = 1
    GetINIT id
'        WriteLog "GetInit"
    GetOffSet id
'        WriteLog "GetOffset"
    LastDate = sLastDate(id) '自分が管理するファイルの最終日時
'        WriteLog "Get LastDate"
'    Debug.Print DTMtoFname(LastDate)
'    LastFilename = KN_Path(id) & "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_Total.txt"
    LastFilename = "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_Total.txt"
    
   ' WriteLog id & ":" & LastFilename
'    kanrihantei 1, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"
    
    
    Dim co As Integer
    Dim tsFile() As String
    co = CheckDataFile(KN_Path(id), tsFile())
    
'        WriteLog "CheckDataFile : " & co
   
   ' WriteLog id & ":" & co
    
    If 0 < co Then
        Call AppendData(id, tsFile())
'        If Exists2(App.Path & "\fSoushin.exe") = True Then
'            Call Shell(App.Path & "\fSoushin.exe", vbNormalFocus)
'        End If
    End If

    id = 2
    GetINIT id
    GetOffSet id
    LastDate = sLastDate(id) '自分が管理するファイルの最終日時
'    Debug.Print DTMtoFname(LastDate)
    LastFilename = "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_TOTAL.TXT"
'kanrihantei id, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"
    
    co = 0
    Erase tsFile
    co = CheckDataFile(KN_Path(id), tsFile())
    
    If 0 < co Then
        Call AppendData(id, tsFile())
'        If Exists2(App.Path & "\fSoushin.exe") = True Then
'            Call Shell(App.Path & "\fSoushin.exe", vbNormalFocus)
'        End If
'    Else
'        GoTo Main9999
    End If

    '2019-06-28　最後に一回だけ送信を呼び出す。
    Call Shell(App.Path & "\fSoushin.exe", vbNormalFocus)

    If Exists2(App.Path & "\kmSoushin.exe") = True Then
        Call Shell(App.Path & "\kmSoushin.exe", vbNormalFocus)
    End If

Main9999:
End Sub

Private Sub GetINIT(id As Integer)
    Dim Fso   As New FileSystemObject
    Dim FsoTS   As TextStream
    
On Error Resume Next
    
    Dim sa As Variant
    Dim sb As Variant
    Dim bf As String
    Dim i As Integer
    Dim j As Integer
    
        Set FsoTS = Fso.OpenTextFile(KN_table(id), ForReading, False, TristateUseDefault)
        'ファイル全体を読み込み
        bf = FsoTS.ReadAll
        'オープンしていたファイルを閉じる
        FsoTS.Close
        Set FsoTS = Nothing
    
    sa = Split(bf, vbCrLf)
    For i = 0 To UBound(sa)
        If sa(i) <> "" Then
            Select Case Left$(sa(i), 1)
            Case ";", ":", "'"
            Case Else
                sb = Split(sa(i), ",")
                j = sb(0)
                Pname(j) = sb(1)
                INIT(j).x = sb(2)
                INIT(j).y = sb(3)
                INIT(j).z = sb(4)
                PnameK(j) = sb(5)
            End Select
        End If
    Next i
   
On Error GoTo 0
   
End Sub

Private Sub GetOffSet(id As Integer)
    Dim Fso   As New FileSystemObject
    Dim FsoTS   As TextStream
    
    Dim sa As Variant
    Dim sb As Variant
    Dim bf As String
    Dim i As Integer
    Dim j As Integer
    
        Set FsoTS = Fso.OpenTextFile(KN_Offset(id), ForReading, False, TristateUseDefault)
        'ファイル全体を読み込み
        bf = FsoTS.ReadAll
        'オープンしていたファイルを閉じる
        FsoTS.Close
        Set FsoTS = Nothing
    
    sa = Split(bf, vbCrLf)
    For i = 0 To UBound(sa)
        If sa(i) <> "" Then
            Select Case Left$(sa(i), 1)
            Case ";", ":", "'"
            Case Else
                sb = Split(sa(i), ",")
                j = sb(0)
'                Pname(j) = sb(1)
                offSET(j).x = sb(2)
                offSET(j).y = sb(3)
                offSET(j).z = sb(4)
            End Select
        End If
    Next i
   
End Sub

Function zahyohenkan(dt() As zahyo)
    Dim a11 As Double
    Dim a12 As Double
    Dim a21 As Double
    Dim a22 As Double
    
    a11 = Cos(kakudo * RAD)
    a12 = Sin(kakudo * RAD)
    a21 = -Sin(kakudo * RAD)
    a22 = Cos(kakudo * RAD)
    
    Dim i As Integer
    Dim x As Double, y As Double
    Dim xx As Double, yy As Double
    
    For i = 1 To UBound(dt)
        x = dt(i).x
        y = dt(i).y
        xx = a11 * x + a12 * y
        yy = a21 * x + a22 * y
        
        dt(i).x = xx
        dt(i).y = yy
    Next i
    
    
End Function



Public Function CheckDataFile(ByVal fdir As String, tFile() As String) As Long
    
    Dim lIndex As Long
    Dim hFolder As Folder
    Dim hFile As File
    Dim Fso As FileSystemObject
    
    Dim FileList() As String
    Dim i As Long
    Dim j As Integer

    Dim ret As String

    Dim tFilename() As String
    Dim aIndex As Long
    aIndex = -1
    
    Set Fso = New FileSystemObject
    Set hFolder = Fso.GetFolder(fdir)
    lIndex = 0
    For Each hFile In hFolder.Files
'    List1.List(lIndex) = hFile.Path
        'Debug.Print hFile.Path
        If Left$(hFile.Name, 1) = "R" And UCase(Right$(hFile.Name, 3)) = "TXT" Then
            lIndex = lIndex + 1
                aIndex = aIndex + 1
                ReDim Preserve tFilename(aIndex) As String
                tFilename(aIndex) = hFile.Name
        End If
    Next hFile
    
    Set Fso = Nothing
    Set hFile = Nothing
    Set hFolder = Nothing

    CheckDataFile = lIndex
    If lIndex = 0 Then Exit Function

        '所得したファイル名をソート
        If -1 < aIndex Then
            s_ShellSort tFilename(), (aIndex)
        End If
        
    For i = 0 To aIndex
        If UCase(LastFilename) < UCase(tFilename(i)) Then
            j = j + 1
            ReDim Preserve tFile(j)
            tFile(j) = tFilename(i)
        End If
    Next i
    CheckDataFile = j

Exit Function

'ファイル検索
'ファイルがあったら、配列にパス名を取得
'fDir : 検索ディレクトリ

    On Error GoTo CheckDataFile9999

    aIndex = -1

    If GetTargetFiles(FileList, fdir, "TXT") Then
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
'    CheckDataFile = aIndex + 1
    
        
    For i = 0 To aIndex
        If LastFilename < tFilename(i) Then
            j = j + 1
            ReDim Preserve tFile(j)
            tFile(j) = tFilename(i)
        End If
    Next i
    CheckDataFile = j
    
    On Error GoTo 0
    Exit Function
    
CheckDataFile9999:
    CheckDataFile = 0
    On Error GoTo 0
End Function


Function FnametoDTM(ByVal st As String) As String
'TSファイル名から日時を生成
'st : TSファイル名 20170826_09
    
On Error GoTo FnametoDTM9999
    st = Replace(st, ".txt", "")
    st = Replace(st, ".TXT", "")
    st = Replace(UCase(st), "_TOTAL", "")
    Dim sst As String
    Dim yy As String
    Dim mm As String
    Dim dd As String
    Dim hh As String
    Dim nn As String
    Dim ss As String
    
    yy = Mid$(st, 2, 4)
    mm = Mid$(st, 6, 2)
    dd = Mid$(st, 8, 2)
    hh = Mid$(st, 10, 2)
    nn = Mid$(st, 12, 2)
    ss = Mid$(st, 14, 2)
    
    sst = Format(DateSerial(yy, mm, dd) + TimeSerial(hh, nn, ss), "yyyy/mm/dd hh:mm:ss")
    
    FnametoDTM = sst
    On Error GoTo 0
Exit Function
FnametoDTM9999:
    FnametoDTM = ""
    On Error GoTo 0
End Function

Function DTMtoFname(st As String) As String
'日時フォーマットからファイル名を生成
'st : 日時フォーマット
    Dim sst As String
    If IsDate(st) = True Then
        sst = Format(st, "yyyymmddhhmmss")
        DTMtoFname = sst
    Else
        DTMtoFname = ""
    End If
End Function

Function DTMtoDname(st As String) As String
'日時フォーマットからディレクトリ名を生成
'st : 日時フォーマット
    Dim sst As String
    If IsDate(st) = True Then
        sst = Format(st, "yyyy-mm")
        DTMtoDname = sst
    Else
        DTMtoDname = ""
    End If
End Function

Function sLastDate(id As Integer) As String
'保存データファイルの最終日時を取得する
' ID : データ番号
' ed : 最終日時

On Error GoTo LastDate9999

    Dim ed As String

    Dim Fso     As New FileSystemObject
    Dim FsoTS   As TextStream

    Dim MaxLine As Long
    'ファイルの末尾から書き込みモードで開きます
    Set FsoTS = Fso.OpenTextFile(oFile(id), ForAppending)
    '現在のファイル ポインターの位置を行番号で取得します
    MaxLine = FsoTS.Line - 1
    FsoTS.Close
    
    Dim i As Integer
    Set FsoTS = Fso.OpenTextFile(oFile(id))
    '指定行まで読み飛ばし（For～Nextでの読み飛ばしの方がDo～Loopより早い）
    For i = 1 To MaxLine - 1
        FsoTS.SkipLine
    Next i
    'ファイルの最後以降は取得しない(最後にvbCrLfがあると１行増えてエラーが発生
    'する場合があるのでvbCrLfだけの行は取得しない)
    If FsoTS.AtEndOfStream = False Then
        '指定行のデータを取得
        ed = FsoTS.ReadLine ' & vbCrLf
    End If
    FsoTS.Close
    Set FsoTS = Nothing

    ed = Mid$(ed, 1, 19)

    sLastDate = ed
    On Error GoTo 0
Exit Function
LastDate9999:
    sLastDate = ""
    On Error GoTo 0
End Function

Sub AppendData(ByVal id As Integer, fNam() As String)
'
    Dim n1 As String
    Dim Fso   As New FileSystemObject
    Dim FsoTS As TextStream
    Dim FsoMS As TextStream
'    Dim FsoHS As TextStream
    Dim FsoSS As TextStream
    Dim bf As String
    Dim wbf As String
    Dim kbf As String
    
    Dim ii As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sa As Variant
    Dim sb As Variant
    
    Dim MDY As String
    ReDim dt(sokutenSu(id)) As zahyo
    ReDim heniDT(sokutenSu(id)) As zahyo
    ReDim heni(sokutenSu(id)) As zahyo
    Dim cc As Integer
    Dim no As Integer
    Dim fx As Boolean
    Dim tID As Integer
    
    Dim wmData(2) As Double
    
'■■■ ADD START 2020/08/20 T.Kimura エラー発生してもプログラムを中断せず処理を継続させる ■■■
    On Error Resume Next
'■■■ ADD END   2020/08/20 T.Kimura エラー発生してもプログラムを中断せず処理を継続させる ■■■
    
    For i = 0 To sokutenSu(id)
        dt(i).x = 999999
        dt(i).y = 999999
        dt(i).z = 999999
        heniDT(i).x = 999999
        heniDT(i).y = 999999
        heniDT(i).z = 999999
        heni(i).x = 999999
        heni(i).y = 999999
        heni(i).z = 999999
    Next i

    If id = 3 Then
        tID = 2
    Else
        tID = id
    End If
    
'    WriteLog (tID) & " - " & App.Path & "\" & oFile(id)

'    WriteLog App.Path & "\" & oFile(id)
    'OpenTextFile(ファイル名, 1, True, -2)  読み込み専用
    'OpenTextFile(ファイル名, 2, True, -2)  書き込みもできる
    'OpenTextFile(ファイル名, 8, True, -2)  追記  OSのデフォルト
    'OpenTextFile(ファイル名, 8, True, -1)        Shift_jis
    Set FsoMS = Fso.OpenTextFile(App.Path & "\" & oFile(id), 8, True, -2)
'    Set FsoHS = Fso.OpenTextFile(App.Path & "\" & heFile, 8, True, -2)

'    WriteLog oFile(id)

    For ii = 1 To UBound(fNam())
    
'    WriteLog KN_Path(tID) & UCase(fNam(ii))
    
        n1 = UCase(fNam(ii))
        If UCase("R" & DTMtoFname(LastDate) & KN_SubName(tID) & "_Total.txt") < n1 Then
            If Fso.FileExists(KN_Path(tID) & n1) = False Then
                Exit Sub
            End If
        
            fx = True
        
            Set FsoTS = Fso.OpenTextFile(KN_Path(tID) & n1, ForReading, False, TristateUseDefault)
            'ファイル全体を読み込み
            bf = FsoTS.ReadAll
            'オープンしていたファイルを閉じる
            FsoTS.Close
            Set FsoTS = Nothing
    
            MDY = FnametoDTM(n1)
            
            sa = Split(bf, vbCrLf)
            For i = 0 To UBound(sa)
                If sa(i) <> "" Then
                    sb = Split(sa(i), ",")
                    For j = 1 To sokutenSu(id)
                        If UCase(sb(1) & sb(2)) = Pname(j) Then
                            If sb(3) = 0 And sb(4) = 0 Then
                                no = j 'sb(2)
                                dt(no).x = sb(14)
                                dt(no).y = sb(16)
                                dt(no).z = sb(18)
                                heniDT(no).x = sb(6)
                                heniDT(no).y = sb(8)
                                heniDT(no).z = sb(10)
'                                Debug.Print Pname(j), j
                            End If
                            Exit For
                        End If
                    Next j
                End If
            Next i
        '    Debug.Print sa(0)
            wbf = MDY
            For i = 1 To sokutenSu(id)
                wbf = wbf & "," & dt(i).x & "," & dt(i).y & "," & dt(i).z
            Next i
            FsoMS.WriteLine wbf
             
            '座標データ保存
            Set FsoSS = Fso.OpenTextFile(SoushinPathZ(id) & Format(MDY, "yyyy-mm-dd_hh-mm-ss") & ".csv", 2, True, -2)
            wbf = MDY
            For i = 1 To sokutenSu(id)
                wbf = wbf & "," & Format(dt(i).x, "0.0000") & "," & Format(dt(i).y, "0.0000") & "," & Format(dt(i).z, "0.0000")
            Next i
            FsoSS.WriteLine (wbf)
            FsoSS.Close
            
            '変位量 (mm)を求める
            For i = 1 To sokutenSu(id)
                If heniDT(i).x = 999999 Then
                    heni(i).x = 999999
                Else
                    heni(i).x = (dt(i).x - INIT(i).x) * 1000# - offSET(i).x
                End If
                If heniDT(i).y = 999999 Then
                    heni(i).y = 999999
                Else
                    heni(i).y = (dt(i).y - INIT(i).y) * 1000# - offSET(i).y
                End If
                If heniDT(i).z = 999999 Then
                    heni(i).z = 999999
                Else
                    heni(i).z = (dt(i).z - INIT(i).z) * 1000# - offSET(i).z
                End If
            Next i
            
            '変位量保存
            Set FsoSS = Fso.OpenTextFile(SoushinPath(id) & Format(MDY, "yyyy-mm-dd_hh-mm-ss") & ".csv", 2, True, -2)
            wbf = MDY
            For i = 1 To sokutenSu(id)
                wbf = wbf & "," & FormatD(heni(i).x, "0.0") & "," & FormatD(heni(i).y, "0.0") & "," & FormatD(heni(i).z, "0.0")
            Next i
            FsoSS.WriteLine (wbf)
            FsoSS.Close
        
            '最終データ保存
            Set FsoSS = Fso.OpenTextFile(App.Path & "\Newest" & id & ".csv", 2, True, -2)
            wbf = MDY
            For i = 1 To sokutenSu(id)
                wbf = wbf & "," & FormatD(heni(i).x, "0.0000") & "," & FormatD(heni(i).y, "0.0000") & "," & FormatD(heni(i).z, "0.0000")
            Next i
            FsoSS.WriteLine (wbf)
            FsoSS.Close
        
        '気象用ファイル作成
        
            MDY = FnametoDTM(n1)
            
            sa = Split(bf, vbCrLf)
            For i = 0 To UBound(sa)
                If sa(i) <> "" Then
                    sb = Split(sa(i), ",")
                        If UCase(sb(1)) = "WM" Then
                                wmData(1) = sb(6)  '温度
                                wmData(2) = sb(8)  '気圧
                                Exit For
                        End If
                End If
            Next i
        '    Debug.Print sa(0)
            
            '気象データ保存
            Set FsoSS = Fso.OpenTextFile(SoushinWM(id) & Format(MDY, "yyyy-mm-dd_hh-mm-ss") & ".csv", 2, True, -2)
            kbf = MDY
            For i = 1 To 2
                kbf = kbf & "," & wmData(i)
            Next i
            FsoSS.WriteLine (kbf)
            FsoSS.Close
        
        
        End If
'            If id <> 2 Then
                DoFileMove KN_Path(tID) & n1, KN_PathBK(tID) & n1
'            End If
    Next ii
    FsoMS.Close
'    FsoHS.Close
    Set FsoMS = Nothing '
'    Set FsoHS = Nothing '
    Set FsoSS = Nothing '0
    
    If fx = True Then
        Call kanrihantei(id, wbf)
    End If
    
End Sub

Public Function FormatD(dt As Double, fmt As String) As String
    If Abs(dt) = 999999 Then
        FormatD = "999999"
    Else
        FormatD = Format(dt, fmt)
    End If
End Function

'2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978
Sub kanrihantei(id As Integer, bf As String)
    Dim sa As Variant
    Dim i As Integer
    
    Dim xd(40) As Double
    Dim yd(40) As Double
    Dim zd(40) As Double
    
    kanriLV = -1
    sa = Split(bf, ",")
    For i = 1 To UBound(sa)
        Select Case (i Mod 3)
        Case 1
            xd((i \ 3) + 1) = sa(i)
        Case 2
            yd((i \ 3) + 1) = sa(i)
        Case 0
            zd((i \ 3) + 0) = sa(i)
        End Select
    Next i
    
    '管理レベルを調べる
    For i = 1 To UBound(sa) \ 3
        If zd(i) <> 999999 Then
            If Not (-kanriV(3) < zd(i) And zd(i) < kanriV(3)) Then
                kanriLV = 3
            ElseIf Not (-kanriV(2) < zd(i) And zd(i) < kanriV(2)) Then
                If kanriLV < 3 Then kanriLV = 2
            ElseIf Not (-kanriV(1) < zd(i) And zd(i) < kanriV(1)) Then
                If kanriLV < 2 Then kanriLV = 1
            Else
                If kanriLV < 1 Then kanriLV = 0
            End If
        End If
    Next i
        
    Dim ret As Integer, sda As String
    sda = sa(0)
    If 0 < kanriLV Then
        If Exists2(App.Path & "\ALARMsw.exe") Then
            Call WriteLog("アラームSW　ON " & kanriLV)
'            Call ALERTfileCK(sda, ret)
            If ret < 0 Then
                Call Shell(App.Path & "\ALARMsw.exe " & kanriLV, vbNormalFocus)
            ElseIf ret < kanriLV Then
                Call Shell(App.Path & "\ALARMsw.exe " & kanriLV, vbNormalFocus)
            Else
                Call Shell(App.Path & "\ALARMsw.exe " & ret, vbNormalFocus)
            End If
        End If
    End If
    
    If kanriLV <= 0 Then
        Exit Sub
    End If
    
    Dim alst As String
    
    alst = "以下の計測データが管理値を超過しました。" & vbCrLf
    alst = alst & vbCrLf & "計測日時：" & sa(0)
    
    '管理レベルを超えた測点を調べる
    For i = 1 To UBound(sa) \ 3
        If zd(i) <> 999999 Then
            If Not (-kanriV(3) < zd(i) And zd(i) < kanriV(3)) Then
'                alst = alst & vbCrLf & "管理レベルⅢ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
'                alst = alst & vbCrLf & "管理レベルⅢ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                alst = alst & vbCrLf & "管理レベルⅢ超過 : " & (PnameK(i)) & " 沈下量 = " & zd(i) & " mm"
            ElseIf Not (-kanriV(2) < zd(i) And zd(i) < kanriV(2)) Then
'                alst = alst & vbCrLf & "管理レベルⅡ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
'                alst = alst & vbCrLf & "管理レベルⅡ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                alst = alst & vbCrLf & "管理レベルⅡ超過 : " & (PnameK(i)) & " 沈下量 = " & zd(i) & " mm"
            ElseIf Not (-kanriV(1) < zd(i) And zd(i) < kanriV(1)) Then
'                alst = alst & vbCrLf & "管理レベルⅠ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
'                alst = alst & vbCrLf & "管理レベルⅠ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                alst = alst & vbCrLf & "管理レベルⅠ超過 : " & (PnameK(i)) & " 沈下量 = " & zd(i) & " mm"
            End If
        End If
    Next i
    
    alst = alst & vbCrLf & "========================================"
    
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\send0000.txt" For Append As #f
    Print #f, alst
    Close #f
    
'    If Exists2(App.Path & "\kmSoushin.exe") = True Then
'        Call Shell(App.Path & "\kmSoushin.exe", vbNormalFocus)
'    End If
    
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

Function SeName(id As Integer, da As String) As String
    Dim p1 As Integer, p2 As Integer
    p1 = InStr(da, "/")
    p2 = InStr(p1 + 1, da, "/")
    If p1 = 0 Or p2 = 0 Then
        SeName = ""
    End If
    
    Dim sYY As String
    Dim sMM As String
    
    sYY = Mid$(da, 1, p1 - 1)
    sMM = Mid$(da, p1 + 1, p2 - p1 - 1)
    SeName = id & "__" & sYY & sMM
End Function

'//// ファイル及びフォルダの有無チェック(有無のみ判定) /////
' True=存在する、False=存在しない
'-----------------------------------------------------------
Public Function Exists2(ByVal strPathName As String) As Boolean
    'strPathName : フルパス名
    '------------------------
    On Error GoTo CheckError
    If (GetAttr(strPathName) And vbDirectory) = vbDirectory Then
        'Debug.Print strPathName & "はフォルダです。"
    Else
        'Debug.Print strPathName & "はファイルです。"
    End If
    Exists2 = True
    Exit Function
     
CheckError:
    'Debug.Print strPathName & "が見つかりません。"
    On Error GoTo 0
End Function

Sub ALERTfileCK(sda As String, ret As Integer)
On Error Resume Next
    ret = -1
    
    Dim da1 As Date
    Dim da2 As Date
    da1 = Format(sda, "yyyy/mm/dd hh:mm:ss")
    Dim f As Integer
    Dim bf As String
    f = FreeFile
    Open ALERTfile For Input As #f
    Line Input #f, bf
    da2 = Format(bf, "yyyy/mm/dd hh:mm:ss")
    Line Input #f, bf
    Close #f
    
    If Format(da2, "yyyy/mm/dd hh:mm:ss") <= Format(da1, "yyyy/mm/dd hh:mm:ss") Then
        ret = bf
    Else
        ret = -1
    End If
    
On Error GoTo 0
End Sub

Sub test()
    Dim Fso   As New FileSystemObject
    Dim FsoTS   As TextStream
    
    Dim sa As Variant
    Dim bf As String
    
        Set FsoTS = Fso.OpenTextFile(App.Path & "\2017-11\CalcZ_1.csv", ForReading, False, TristateUseDefault)
        'ファイル全体を読み込み
        bf = FsoTS.ReadAll
        'オープンしていたファイルを閉じる
        FsoTS.Close
        Set FsoTS = Nothing
    
    sa = Split(bf, vbCrLf)
    
    Dim i As Integer
    For i = 0 To UBound(sa) - 1
        Debug.Print sa(i)
    Next i
    

End Sub

Sub DoFileMove(sp As String, dp As String)
'sp:元
'dp:先
On Error Resume Next

    Dim Fso As New FileSystemObject
    
    Dim ssp As String
    Dim sdp As String
    ssp = Fso.GetAbsolutePathName(sp)
    sdp = Fso.GetAbsolutePathName(dp)
    
    Dim dDirectory As String
    Dim fNam As String
    Dim pa As String
    dDirectory = Fso.GetParentFolderName(sdp) '; // "C:\\data" が返る
    fNam = Fso.GetFileName(sdp)
    pa = FnametoDTM(fNam)
    
    MakeDirectory dDirectory & Format(pa, "\\yyyy\\mm\\dd")
    sdp = dDirectory & Format(pa, "\\yyyy\\mm\\dd") & "\" & fNam
    Fso.CopyFile ssp, sdp, True
    Fso.DeleteFile ssp, True
    
On Error GoTo 0
End Sub

Public Sub MakeDirectory(ByVal Path As String)
    '深い階層のディレクトリまで作成
    SHCreateDirectoryEx 0&, StrPtr(Path), 0&
End Sub

'以下 2017年11月22日 追加
Private Function NameCHG(st As String) As String
    Dim st0 As String
    Dim st1 As String
    Dim st2 As String
    st1 = Left$(st, 1)
    st0 = st1 & "-"
    st2 = Right$(st, Len(st) - 1)
    FindNumberRegExp st2, st1
    NameCHG = st0 & st1
End Function

'// 引数1：対象文字列
'// 引数2：検索結果
Private Sub FindNumberRegExp(s As String, Result As String)
    If InStr(s, "0") = 0 Then
        Result = s
        Exit Sub
    End If
    
    Dim reg     As New RegExp       '// 正規表現クラスオブジェクト
    
    '// 検索条件＝数字を抽出
    reg.Pattern = "[0-9]"
    '// 文字列の最後まで検索する
    reg.Global = True
    '// 指定セルの数字以外の文字を空文字に置き換える
    Result = reg.Replace(s, "")
End Sub
