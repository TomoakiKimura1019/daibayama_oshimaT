Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Text

Imports System.Text.RegularExpressions

Module modMain

    Private Const RAD As Double = 3.14159265358979 / 180.0#

    Dim oPath As String 'TSのデータ格納Path
    Dim oFile(3) As String '自分で管理するのデータファイル名
    Dim oFileA As String 'TSのデータファイル名に付く、日時以外の文字列
    Dim tFile As String
    Dim heFile As String 'TSのデータから変位に
    Dim fUPDATE As Boolean

    Dim LastFilename As String

    Dim LastDate As String

    Structure zahyo
        Dim id As Integer
        Dim x As Double
        Dim y As Double
        Dim z As Double
    End Structure

    Dim kakudo As Double '座標回転角度 DEG

    Dim Pname(14) As String '測点名称
    Dim INIT(14) As zahyo '初期座標 元
    Dim dINIT(14) As zahyo '初期座標 回転後
    Dim offSET(14) As zahyo '変位の補正量 mm


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

    Public GroupName(3) As String

    Public cuDir As String

    Public Sub Main()

        cuDir = My.Application.Info.DirectoryPath
        'test
        cuDir = "C:\work\KN改"
        Dim ini As New IniFile(cuDir & "\TSdata.ini")

        Dim i As Integer
        Dim j As Integer = 0

        KN_Path(1) = ini("system", "KN_Path1")
        KN_Path(2) = ini("system", "KN_Path2")
        If Right(KN_Path(1), 1) <> "\" Then KN_Path(1) = KN_Path(1) & "\"
        If Right(KN_Path(2), 1) <> "\" Then KN_Path(2) = KN_Path(2) & "\"

        KN_PathBK(1) = ini("system", "KN_MovePath1")
        KN_PathBK(2) = ini("system", "KN_MovePath2")
        If Right(KN_PathBK(1), 1) <> "\" Then KN_PathBK(1) = KN_PathBK(1) & "\"
        If Right(KN_PathBK(2), 1) <> "\" Then KN_PathBK(2) = KN_PathBK(2) & "\"

        KN_table(1) = ini("system", "KN_table1")
        KN_table(2) = ini("system", "KN_table2")
        KN_table(3) = ini("system", "KN_table3")

        KN_Offset(1) = ini("system", "KN_Offset1")
        KN_Offset(2) = ini("system", "KN_Offset2")
        KN_Offset(3) = ini("system", "KN_Offset3")

        SoushinPath(1) = ini("system", "SendPath1")
        SoushinPath(2) = ini("system", "SendPath2")
        SoushinPath(3) = ini("system", "SendPath3")
        If Right(SoushinPath(1), 1) <> "\" Then SoushinPath(1) = SoushinPath(1) & "\"
        If Right(SoushinPath(2), 1) <> "\" Then SoushinPath(2) = SoushinPath(2) & "\"
        If Right(SoushinPath(3), 1) <> "\" Then SoushinPath(3) = SoushinPath(3) & "\"

        SoushinPathZ(1) = ini("system", "SendPath1z")
        SoushinPathZ(2) = ini("system", "SendPath2z")
        SoushinPathZ(3) = ini("system", "SendPath3z")
        If Right(SoushinPathZ(1), 1) <> "\" Then SoushinPathZ(1) = SoushinPathZ(1) & "\"
        If Right(SoushinPathZ(2), 1) <> "\" Then SoushinPathZ(2) = SoushinPathZ(2) & "\"
        If Right(SoushinPathZ(3), 1) <> "\" Then SoushinPathZ(3) = SoushinPathZ(3) & "\"

        GroupName(1) = ini("Group", "Name1")
        GroupName(2) = ini("Group", "Name2")
        GroupName(3) = ini("Group", "Name3")

        oPath = ini("system", "oPath")
        oFile(1) = ini("system", "oFile1")
        oFile(2) = ini("system", "oFile2")
        oFile(3) = ini("system", "oFile3")

        heFile = ini("system", "hFile")
        oFileA = ini("system", "oFileA")
        ALERTfile = ini("system", "ALERTfile")

        sokutenSu(1) = 14
        sokutenSu(2) = 13
        sokutenSu(3) = 9

        '1と2を入れ換えている
        KN_SubName(1) = "RAIL02"
        KN_SubName(2) = "RAIL01"

        '            Call ALERTfileCK("2017/09/19 20:00:00", j)

        For i = 1 To 3
            kanriV(i) = CDbl(ini("kanri", "Vkanri" & i))
        Next i


        '"2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978

        LOGFILE = "TDdata.log"


        Dim st As String = ""
        Dim id As Integer

        id = 1
        GetINIT(id)
        GetOffSet(id)
        LastDate = sLastDate(id) '自分が管理するファイルの最終日時
        '    Debug.Print DTMtoFname(LastDate)
        '    LastFilename = KN_Path(id) & "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_Total.txt"
        LastFilename = "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_Total.txt"

        ' WriteLog id & ":" & LastFilename
        '    kanrihantei 1, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"


        Dim co As Integer
        Dim tsFile() As String = {}
        co = CheckDataFile(KN_Path(id), tsFile)

        ' WriteLog id & ":" & co

        If 0 < co Then
            Call AppendData(id, tsFile)
            If Exists2(cuDir & "\fSoushin.exe") = True Then
                Call Shell(cuDir & "\fSoushin.exe", AppWinStyle.NormalFocus)
            End If
        End If

        id = 2
        GetINIT(id)
        GetOffSet(id)
        LastDate = sLastDate(id) '自分が管理するファイルの最終日時
        '    Debug.Print DTMtoFname(LastDate)
        LastFilename = "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_TOTAL.TXT"
        'kanrihantei id, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"

        co = 0
        Erase tsFile
        co = CheckDataFile(KN_Path(id), tsFile)

        If 0 < co Then
            Call AppendData(id, tsFile)
            If Exists2(cuDir & "\fSoushin.exe") = True Then
                Call Shell(cuDir & "\fSoushin.exe", AppWinStyle.NormalFocus)
            End If
        Else
            GoTo Main9999
        End If

        id = 3
        GetINIT(id)
        GetOffSet(id)
        LastDate = sLastDate(id) '自分が管理するファイルの最終日時
        '    Debug.Print DTMtoFname(LastDate)
        'ID=2 と ID=3 は同じファイルをみる
        LastFilename = "R" & DTMtoFname(LastDate) & KN_SubName(2) & "_TOTAL.TXT"
        'kanrihantei id, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"

        co = 0
        Erase tsFile
        co = CheckDataFile(KN_Path(2), tsFile)

        If 0 < co Then
            Call AppendData(id, tsFile)
            If Exists2(cuDir & "\fSoushin.exe") = True Then
                Call Shell(cuDir & "\fSoushin.exe", AppWinStyle.NormalFocus)
            End If
        End If

Main9999:
    End Sub

    Private Sub GetINIT(ByRef id As Integer)

        Dim sa() As String
        Dim sb() As String
        Dim bf As String = ""
        Dim i As Integer
        Dim j As Integer

        Try
            Using sr As New StreamReader( _
                KN_table(id), _
                Encoding.GetEncoding("Shift_JIS"))
                bf = sr.ReadToEnd()
                sr.Close()
            End Using

            'Console.Write(bf)

            sa = Split(bf, Chr(&HD) & Chr(&HA))
            For i = 0 To UBound(sa)
                If sa(i) <> "" Then
                    Select Case Left(sa(i), 1)
                        Case ";", ":", "'"
                        Case Else
                            sb = Split(sa(i), ",")
                            j = CInt(sb(0))
                            Pname(j) = sb(1)
                            INIT(j).x = CDbl(sb(2))
                            INIT(j).y = CDbl(sb(3))
                            INIT(j).z = CDbl(sb(4))
                    End Select
                End If
            Next i

        Catch exception As Exception
            'Console.WriteLine(exception.Message)
        End Try

    End Sub

    Private Sub GetOffSet(ByRef id As Integer)
        Dim sa() As String
        Dim sb() As String
        Dim bf As String
        Dim i As Integer
        Dim j As Integer

        Try
            Using sr As New StreamReader( _
                KN_Offset(id), _
                Encoding.GetEncoding("Shift_JIS"))
                bf = sr.ReadToEnd()
                sr.Close()
            End Using

            sa = Split(bf, Chr(&HD) & Chr(&HA))
            For i = 0 To UBound(sa)
                If sa(i) <> "" Then
                    Select Case Left(sa(i), 1)
                        Case ";", ":", "'"
                        Case Else
                            sb = Split(sa(i), ",")
                            j = CInt(sb(0))
                            '                Pname(j) = sb(1)
                            offSET(j).x = CDbl(sb(2))
                            offSET(j).y = CDbl(sb(3))
                            offSET(j).z = CDbl(sb(4))
                    End Select
                End If
            Next i

        Catch exception As Exception
            'Console.WriteLine(exception.Message)
        End Try
    End Sub

    Sub zahyohenkan(ByRef dt() As zahyo)
        Dim a11 As Double
        Dim a12 As Double
        Dim a21 As Double
        Dim a22 As Double

        a11 = System.Math.Cos(kakudo * RAD)
        a12 = System.Math.Sin(kakudo * RAD)
        a21 = -System.Math.Sin(kakudo * RAD)
        a22 = System.Math.Cos(kakudo * RAD)

        Dim i As Integer
        Dim x, y As Double
        Dim xx, yy As Double

        For i = 1 To UBound(dt)
            x = dt(i).x
            y = dt(i).y
            xx = a11 * x + a12 * y
            yy = a21 * x + a22 * y

            dt(i).x = xx
            dt(i).y = yy
        Next i


    End Sub


    Public Function CheckDataFile(ByVal fdir As String, ByRef tFile() As String) As Integer

        Dim lIndex As Integer

        Dim i As Integer
        Dim j As Integer

        Dim ret As String = ""

        Dim tFilename() As String = {}
        Dim aIndex As Integer
        aIndex = -1

        lIndex = 0

        For Each filepath As String In Directory.GetFiles(fdir, "*.txt", SearchOption.TopDirectoryOnly)
            If Left(FindFileName(filepath), 1) = "R" And UCase(Right(FindFileName(filepath), 3)) = "TXT" Then
                lIndex = lIndex + 1
                aIndex = aIndex + 1
                ReDim Preserve tFilename(aIndex)
                tFilename(aIndex) = FindFileName(filepath)
            End If
            '            Console.WriteLine(filepath)
        Next

        CheckDataFile = lIndex
        If lIndex = 0 Then
            CheckDataFile = 0
            Exit Function
        End If

        '所得したファイル名をソート
        If -1 < aIndex Then
            s_ShellSort(tFilename, (aIndex))
        End If

        For i = 0 To aIndex
            If UCase(LastFilename) < UCase(tFilename(i)) Then
                j = j + 1
                ReDim Preserve tFile(j)
                tFile(j) = tFilename(i)
            End If
        Next i
        CheckDataFile = j

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

        yy = Mid(st, 2, 4)
        mm = Mid(st, 6, 2)
        dd = Mid(st, 8, 2)
        hh = Mid(st, 10, 2)
        nn = Mid(st, 12, 2)
        ss = Mid(st, 14, 2)

        Dim dt As New DateTime(CInt(yy), CInt(mm), CInt(dd), CInt(hh), 0, 0, DateTimeKind.Local)
        sst = dt.ToString("yyyy/MM/dd HH:mm:ss")

        FnametoDTM = sst
        On Error GoTo 0
        Exit Function
FnametoDTM9999:
        FnametoDTM = ""
        On Error GoTo 0
    End Function

    Function DTMtoFname(ByRef st As String) As String
        '日時フォーマットからファイル名を生成
        'st : 日時フォーマット
        Dim dt As DateTime
        Dim sst As String
        If DateTime.TryParse(st, dt) = True Then
            sst = dt.ToString("yyyyMMddHHmmss")
            DTMtoFname = sst
        Else
            DTMtoFname = ""
        End If
    End Function

    Function DTMtoDname(ByRef st As String) As String
        '日時フォーマットからディレクトリ名を生成
        'st : 日時フォーマット
        Dim dt As DateTime
        Dim sst As String
        If DateTime.TryParse(st, dt) = True Then
            sst = dt.ToString("yyyy-MM")
            DTMtoDname = sst
        Else
            DTMtoDname = ""
        End If
    End Function

    Function sLastDate(ByRef id As Integer) As String
        '保存データファイルの最終日時を取得する
        ' ID : データ番号
        ' ed : 最終日時

        On Error GoTo LastDate9999

        Dim nm As String = oFile(id)
        Dim ed As String = ""

        Dim fi As New System.IO.FileInfo(nm)
        'ファイルのサイズを取得
        Dim l As Long = fi.Length
        Dim sl As Long = 0
        Dim sp As Long = 0
        sl = l
        Do
            sl = sl \ 2
            If sl < 1024 Then
                sp = l - sl
                Dim fs As FileStream
                Dim sr As StreamReader
                Dim buf(2) As Char

                fs = New FileStream(nm, FileMode.Open, FileAccess.Read)
                sr = New StreamReader(fs)
                fs.Seek(sp, SeekOrigin.Begin)
                'sr.ReadBlock(buf, 0, buf.Length)

                Do While -1 < sr.Peek
                    ed = sr.ReadLine()
                Loop

                sr.Close()
                fs.Close()
                ed = Mid(ed, 1, 19)
                sLastDate = ed
                Exit Do
            End If
        Loop
        On Error GoTo 0
        Exit Function

LastDate9999:
        sLastDate = ""
        On Error GoTo 0
    End Function

    Sub AppendData(ByVal id As Integer, ByRef fNam() As String)

        Dim n1 As String
        Dim bf As String
        Dim wbf As String = ""

        Dim ii As Integer
        Dim i As Integer
        Dim j As Integer
        Dim sa As Object
        Dim sb As Object

        Dim MDY As String
        Dim dt(14) As zahyo
        Dim heniDT(14) As zahyo
        Dim heni(14) As zahyo
        Dim cc As Integer = 0
        Dim no As Integer
        Dim fx As Boolean
        Dim tID As Integer

        For i = 0 To 14
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

        Dim swMS As New System.IO.StreamWriter(oFile(id), _
           True, _
           System.Text.Encoding.GetEncoding("shift_jis"))

        For ii = 1 To UBound(fNam)

            '    WriteLog KN_Path(tID) & UCase(fNam(ii))

            n1 = UCase(fNam(ii))
            If UCase("R" & DTMtoFname(LastDate) & KN_SubName(tID) & "_Total.txt") < n1 Then
                If System.IO.File.Exists(KN_Path(tID) & n1) = False Then
                    Exit Sub
                End If

                fx = True

                Dim swTS As New System.IO.StreamReader(KN_Path(tID) & n1, _
                   System.Text.Encoding.GetEncoding("shift_jis"))
                'ファイル全体を読み込み
                bf = swTS.ReadToEnd()
                'オープンしていたファイルを閉じる
                swTS.Close()

                MDY = FnametoDTM(n1)

                sa = Split(bf, Chr(&HD) & Chr(&HA))
                For i = 0 To UBound(sa)
                    If sa(i) <> "" Then
                        sb = Split(sa(i), ",")
                        For j = 1 To sokutenSu(id)
                            If UCase(sb(1) & sb(2)) = Pname(j) Then
                                If sb(3) = 0 And sb(4) = 0 Then
                                    no = j 'sb(2)
                                    dt(no).x = CDbl(sb(14))
                                    dt(no).y = CDbl(sb(16))
                                    dt(no).z = CDbl(sb(18))
                                    heniDT(no).x = CDbl(sb(6))
                                    heniDT(no).y = CDbl(sb(8))
                                    heniDT(no).z = CDbl(sb(10))
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
                swMS.WriteLine(wbf)

                Dim dte As DateTime

                DateTime.TryParse(MDY, dte)

                Dim swSS As New System.IO.StreamWriter(SoushinPathZ(id) & dte.ToString("yyyy-MM-dd_HH-mm-ss") & ".csv", _
                    False, _
                   System.Text.Encoding.GetEncoding("shift_jis"))

                wbf = MDY
                For i = 1 To sokutenSu(id)
                    wbf = wbf & "," & dt(i).x.ToString("0.0000") & "," & dt(i).y.ToString("0.0000") & "," & dt(i).z.ToString("0.0000")
                Next i
                swSS.WriteLine((wbf))
                swSS.Close()
                '変位量 (mm)
                For i = 1 To sokutenSu(id)
                    If heniDT(i).x = 999999 Then
                        heni(i).x = 999999
                    Else
                        heni(i).x = (heniDT(i).x - INIT(i).x) - offSET(i).x
                    End If
                    If heniDT(i).y = 999999 Then
                        heni(i).y = 999999
                    Else
                        heni(i).y = (heniDT(i).y - INIT(i).y) - offSET(i).y
                    End If
                    If heniDT(i).z = 999999 Then
                        heni(i).z = 999999
                    Else
                        heni(i).z = (heniDT(i).z - INIT(i).z) - offSET(i).z
                    End If
                Next i

                Dim swSS2 As New System.IO.StreamWriter(SoushinPath(id) & dte.ToString("yyyy-MM-dd_HH-mm-ss") & ".csv", _
                    False, _
                   System.Text.Encoding.GetEncoding("shift_jis"))

                wbf = MDY
                For i = 1 To sokutenSu(id)
                    wbf = wbf & "," & FormatD(heni(i).x, "0.0000") & "," & FormatD(heni(i).y, "0.0000") & "," & FormatD(heni(i).z, "0.0000")
                Next i
                swSS2.WriteLine((wbf))
                swSS2.Close()

                Dim swSS3 As New System.IO.StreamWriter(cuDir & "\Newest" & id & ".csv", _
                    False, _
                   System.Text.Encoding.GetEncoding("shift_jis"))
                wbf = MDY
                For i = 1 To sokutenSu(id)
                    wbf = wbf & "," & FormatD(heni(i).x, "0.0000") & "," & FormatD(heni(i).y, "0.0000") & "," & FormatD(heni(i).z, "0.0000")
                Next i
                swSS3.WriteLine((wbf))
                swSS3.Close()

            End If
            If id <> 2 Then
                DoFileMove(KN_Path(tID) & n1, KN_PathBK(tID) & n1)
            End If
        Next ii
        '閉じる
        swMS.Close()

        If fx = True Then
            Call kanrihantei(id, wbf)
        End If

    End Sub

    Public Function FormatD(ByRef dt As Double, ByRef fmt As String) As String
        If System.Math.Abs(dt) = 999999 Then
            FormatD = "999999"
        Else
            FormatD = dt.ToString(fmt)
        End If
    End Function

    '2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978
    Sub kanrihantei(ByRef id As Integer, ByRef bf As String)
        Dim sa() As String
        Dim i As Integer

        Dim xd(14) As Double
        Dim yd(14) As Double
        Dim zd(14) As Double

        kanriLV = -1
        sa = Split(bf, ",")
        For i = 1 To UBound(sa)
            Select Case (i Mod 3)
                Case 1
                    xd((i \ 3) + 1) = CDbl(sa(i))
                Case 2
                    yd((i \ 3) + 1) = CDbl(sa(i))
                Case 0
                    zd((i \ 3) + 0) = CDbl(sa(i))
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

        Dim ret As Integer
        Dim sda As String
        sda = sa(0)
        If 0 < kanriLV Then
            If Exists2(cuDir & "\ALARMsw.exe") Then
                Call WriteLog("アラームSW　ON " & kanriLV)
                '            Call ALERTfileCK(sda, ret)
                If ret < 0 Then
                    Call Shell(cuDir & "\ALARMsw.exe " & kanriLV, AppWinStyle.NormalFocus)
                ElseIf ret < kanriLV Then
                    Call Shell(cuDir & "\ALARMsw.exe " & kanriLV, AppWinStyle.NormalFocus)
                Else
                    Call Shell(cuDir & "\ALARMsw.exe " & ret, AppWinStyle.NormalFocus)
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
                    alst = alst & vbCrLf & "管理レベルⅢ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                ElseIf Not (-kanriV(2) < zd(i) And zd(i) < kanriV(2)) Then
                    '                alst = alst & vbCrLf & "管理レベルⅡ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
                    alst = alst & vbCrLf & "管理レベルⅡ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                ElseIf Not (-kanriV(1) < zd(i) And zd(i) < kanriV(1)) Then
                    '                alst = alst & vbCrLf & "管理レベルⅠ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
                    alst = alst & vbCrLf & "管理レベルⅠ超過 : " & NameCHG(Pname(i)) & " 沈下量 = " & zd(i) & " mm"
                End If
            End If
        Next i

        alst = alst & vbCrLf & "========================================"

        Dim f As Integer
        f = FreeFile()
        FileOpen(f, cuDir & "\send0000.txt", OpenMode.Append)
        PrintLine(f, alst)
        FileClose(f)

        If Exists2(cuDir & "\kmSoushin.exe") = True Then
            Call Shell(cuDir & "\kmSoushin.exe", AppWinStyle.NormalFocus)
        End If

    End Sub

    Public Sub WriteLog(ByRef st As String)
        'st 説明文
        Dim f As Integer

        On Error GoTo WriteLog9999

        f = FreeFile()
        FileOpen(f, cuDir & "\" & LOGFILE, OpenMode.Append)
        Print(f, Now.ToString("yyyy/MM/dd HH:mm:ss") & " : ")
        PrintLine(f, st)
        FileClose(f)

WriteLog9999:
        On Error GoTo 0
    End Sub

    Function SeName(ByRef id As Integer, ByRef da As String) As String
        Dim p1, p2 As Integer
        p1 = InStr(da, "/")
        p2 = InStr(p1 + 1, da, "/")
        If p1 = 0 Or p2 = 0 Then
            SeName = ""
        End If

        Dim sYY As String
        Dim sMM As String

        sYY = Mid(da, 1, p1 - 1)
        sMM = Mid(da, p1 + 1, p2 - p1 - 1)
        SeName = id & "__" & sYY & sMM
    End Function

    '//// ファイル及びフォルダの有無チェック(有無のみ判定) /////
    ' True=存在する、False=存在しない
    '-----------------------------------------------------------
    Public Function Exists2(ByVal strPathName As String) As Boolean
        'strPathName : フルパス名
        '------------------------
        On Error GoTo CheckError
        If (GetAttr(strPathName) And FileAttribute.Directory) = FileAttribute.Directory Then
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

    Sub test()
    End Sub

    Sub DoFileMove(ByRef sp As String, ByRef dp As String)
        'sp:元
        'dp:先
        Dim Fso As New Scripting.FileSystemObject

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

        MakeDirectory(dDirectory & VB6.Format(pa, "\yyyy\MM\dd"))
        sdp = dDirectory & VB6.Format(pa, "\yyyy\MM\dd") & "\" & fNam
        Fso.CopyFile(ssp, sdp, True)
        Fso.DeleteFile(ssp, True)

    End Sub

    Public Sub MakeDirectory(ByVal sPath As String)
        '深い階層のディレクトリまで作成
        System.IO.Directory.CreateDirectory(sPath)
    End Sub

    '以下 2017年11月22日 追加
    Private Function NameCHG(ByRef st As String) As String
        Dim st0 As String
        Dim st1 As String
        Dim st2 As String
        st1 = Left(st, 1)
        st0 = st1 & "-"
        st2 = Right(st, Len(st) - 1)
        FindNumberRegExp(st2, st1)
        NameCHG = st0 & st1
    End Function

    '// 引数1：対象文字列
    '// 引数2：検索結果
    Private Sub FindNumberRegExp(ByRef s As String, ByRef Result As String)
        If InStr(s, "0") = 0 Then
            Result = s
            Exit Sub
        End If
        'Dim result As Boolean = Regex.IsMatch("{検査対象文字列}", "{正規表現パターン}")
        Result = Regex.IsMatch(s, "[0-9]")

    End Sub
End Module