Attribute VB_Name = "MainModule"
Option Explicit

'---------------------------------------------------------------
'自動計測
'---------------------------------------------------------------
'作成履歴
'   2000.06(寺島） 苫東コールセンター TDS-303
'   2000.12(寺島） 磐城 地すべり観測システム TDS-302
'   2000.03(寺島） 札幌ドーム屋根変位計測システム
'---------------------------------------------------------------
'修正履歴
'---------------------------------------------------------------
Public Const PI As Double = 3.1415926535898

'定数
Public Const MaxKOU As Integer = 4     '計測項目数
Public Const Maxdan As Integer = 1     '計測断面数
Public Const Maxtype As Integer = 2    '計測項目の種類数
Public Const TDS_CH As Integer = 1000   'ＴＤＳの最大チャンネル数
Public Const KeijiMAX As Integer = 1   '経時変化図の最大表示数
Public Const MAX_CH  As Integer = 27   '      〃     のデータ数
Public Const REC_LEN As Integer = 291  'データファイルのバイト数 27*10+21

'環境ファイル名
Public Const DATA_DAT   As String = "master.dat" '計測ﾃﾞｰﾀ
Public Const CTABLE_DAT As String = "ctable.dat" '環境設定ﾃﾞｰﾀ
Public Const INTV_DAT   As String = "INTV.dat"

'フォルダ名
Public CurrentDir As String   'keisoku.exeがあるフォルダ
Public Type Kei1
    Tabl_path As String       '共通ファイルフォルダ名(danmen.dat ctable.dat kou.dat)
    keihou_path As String     'keihou.1 keihou.2を出力するフォルダ
    Data_path As String       'データフォルダ名
    BackupPath As String
    STsec As Integer
    STti As String * 40
    Menu(7) As String * 40
End Type
Public KEISOKU As Kei1

'ＴＤＳ読み込みデータフォーマット（1行分）
Type dt
    CH As String * 3  'チャンネル番号
    dt As String * 10 '測定値
End Type

'ＴＤＳ読み込みデータフォーマット（1回分）
Type TDSdt
    Mdate As String * 19     '計測日時
    tDATA(0 To TDS_CH) As dt '測定値
End Type

'データ平均用ＴＤＳ読み込みデータフォーマット（1回分）
Type TDSdt3
    Mdate As String * 19
    CH(TDS_CH) As Integer
    tDATA(TDS_CH) As String * 8
End Type

Public dt1(1 To TDS_CH) As Double       'TDSデータ

'項目設定
Public Type kouSet
    No    As Integer       '項目番号
    KIND  As Integer       '種類番号
    TI1   As String * 20   '項目名
    TI2   As String * 20   '種類名
    Yt    As String * 10   '経時変化図Ｙ軸タイトル
    Yu    As String * 10   '単位
    dec   As Integer       'データ表示時の小数点の四捨五入位置
    max   As Integer       '項目数
    Kmax  As Double        '自動計測経時変化図 スケール最大値
    Kmin  As Double        '        〃            〃  最小値
    KBUN  As Integer       '        〃            〃  分割数
End Type
Public kou(MaxKOU, Maxtype) As kouSet

''環境設定
'Public Type table1
'    dan   As Integer     '断面番号
'    kou   As Integer     '項目番号
'    ten   As Integer     '測点番号
'    FLD   As Integer     'データファイルフィールド位置
'    ch    As Integer     '計測チャンネル
'    Syo   As Double      '初期値
'    Kei   As Double      '係数
'    deep  As double      '深度
'    HAN   As String * 8  '凡例
'    keiji As Integer     '経時変化図作図測点 1.する 0.しない
'End Type
'Public tbl(Maxkou, Maxdan, 10) As table1

'final.dat→データファイルの変換用
Public Type table2
    dan  As Integer
    kou  As Integer
    FLD  As Integer
    CH   As Integer
    Syo  As Double
    Kei  As Double
    Name As String
    keta As Integer
End Type
Public TDSTbl(0 To TDS_CH) As table2
'Public TblCH(0 To TDS_CH) As Integer
'Public TblDAN(0 To TDS_CH) As Integer
'Public TblKOU(0 To TDS_CH) As Integer

'計測インターバル設定
Public Keisoku_Time As Date        '次回計測
Public KE_intv As Date             '計測インターバル
Public Keisoku_TimeZ As Date      '前回計測
Public Lebel_time(4, 24) As Date   'レベル別計測時間
Public Lebel_intv(4) As Integer    'レベル別計測インターバル

'記録インターバル設定
Public kiroku_Time As Date       '次回記録
Public kiroku_intv As Date       '記録インターバル
Public kiroku_TimeZ As Date      '前回記録

'管理値設定
Public keihou_L As Integer    '警報設定管理値１
'Public keihou_L1 As Boolean    '警報設定管理値１
'Public keihou_L2 As Boolean    '警報設定管理値２
'Public keihou_L3 As Boolean    '警報設定管理値３
'Public keihou_L4 As Boolean    '警報設定管理値４

Public Type Kanri1
    Lebel1(4) As Double  '管理値
    Lebel2(4) As Double  '管理値
    TI1(4) As String * 8
    TI2(4) As String * 12
    Hday(4) As Integer  '伸縮計（変位速度の日）
    keihouSW As Integer '警報設定  1=True 0=False
End Type
Public Kanri(MaxKOU, Maxdan) As Kanri1  '警報設定  1=True 0=False
'Public Keihou(4, 4) As Integer '警報設定  1=True 0=False

'フォームサイズ変更時のコマンドボタンのフォームの枠からの位置
Public Const mintControlMargin = 60

Public ivCode '0:計測、1:記録

Public mdy As Date      '計測日時

Public LogDir As String
Public LOGFILE As String

Public KeisokuCO As Integer

Public g_kankyoPath As String

Public KNpath(2) As String

'---------------------------------------------------------------------
Public Sub Main()
    Dim f As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ret As Integer
    Dim SS1 As String
    Dim t1 As Date
    Dim STtime As Date
    
    If App.PrevInstance = True Then
        MsgBox "既に起動しています。", vbCritical, "起動エラー"
        End
    End If
  
 '   Day2Dir "2017/06/15", RetString
  
 '  On Error GoTo TRAP_Main
    
'    ChDrive App.Path
'    ChDir App.Path
    SetCurrentDirectory App.Path


    CurrentDir = App.Path
    If Right$(CurrentDir, 1) = "\" Then Else CurrentDir = CurrentDir & "\"

    
    'ファイル読み込み
    KNpath(1) = GetIni("フォルダ名", "KNpath1", CurrentDir & "計測設定.ini")
    KNpath(2) = GetIni("フォルダ名", "KNpath2", CurrentDir & "計測設定.ini")
    
    Dim sKE_intv As String
    Keisoku_TimeZ = CDate(GetIni("計測時間", "前回計測時間", CurrentDir & "計測設定.ini"))
    Keisoku_Time = CDate(GetIni("計測時間", "次回計測時間", CurrentDir & "計測設定.ini"))
    sKE_intv = (GetIni("計測時間", "計測インターバル", CurrentDir & "計測設定.ini"))
    Call strTOtime(sKE_intv, KE_intv)
    If KE_intv = 0 Then
        KE_intv = #1:00:00 AM#
        Call WriteIni("計測時間", "計測インターバル", Format$(KE_intv, "hh:nn:ss"), CurrentDir & "計測設定.ini")
    End If
    
    kiroku_TimeZ = CDate(GetIni("計測時間", "前回記録時間", CurrentDir & "計測設定.ini"))
    kiroku_Time = CDate(GetIni("計測時間", "次回記録時間", CurrentDir & "計測設定.ini"))
    sKE_intv = (GetIni("計測時間", "記録インターバル", CurrentDir & "計測設定.ini"))
    Call strTOtime(sKE_intv, kiroku_intv)
    If kiroku_intv = 0 Then
        kiroku_intv = #1:00:00 AM#
        Call WriteIni("計測時間", "記録インターバル", Format$(KE_intv, "hh:nn:ss"), CurrentDir & "計測設定.ini")
    End If
    
TusinStart:
    '----------
    
    MainForm.Show
Exit Sub

TRAP_Main:
    
    
    f = FreeFile
    Open CurrentDir & "PRG-err.dat" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : (" & Err.Number & ") " & Err.Description
    Close #f
    
    Close
End Sub

'Public Sub Ktime_ck()
'    '計測時間再設定
'    If DateDiff("s", Keisoku_Time, Now) >= 0 Then  'If nt < Now Then
'        Keisoku_Time = T_ajt(Keisoku_TimeZ, KE_intv)
'        MainForm.xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'    End If
'End Sub

Public Function T_ajt(ByVal d As Date, ByVal T As Date) As Date
'次回計測日時の調整
    Dim nt As Date
    Dim t1 As Date
    Dim t2 As Date
 
    nt = d + T
    t1 = Now
    If DateDiff("s", nt, Now) > 0 Then   'If nt < Now Then
        nt = Date + TimeSerial(Hour(Now), 0, 0)
        Do
            nt = nt + T
            t2 = Now
            If DateAdd("s", 15, t1) < t2 Then Exit Do
        Loop Until DateDiff("s", Now, nt) > 0   'nt > Now
    End If
    T_ajt = nt
End Function

Public Sub ShellSort(Aray() As Date, Num As Integer)
   Dim Span As Integer
   Dim i As Integer
   Dim j As Integer
   Dim TMP As String

   Span = Num \ 2
   Do While Span > 0
      For i = Span To Num - 1
         j% = i% - Span + 1
         For j = (i - Span + 1) To 1 Step -Span
            If Aray(j) >= Aray(j + Span) Then Exit For
            ' 順番の異なる配列要素を入れ替えます.
            TMP = Aray(j)
            Aray(j) = Aray(j + Span)
            Aray(j + Span) = TMP
         Next j
      Next i
      Span = Span \ 2
   Loop
End Sub

Public Sub IntvWrite()
    Call WriteIni("計測時間", "前回計測時間", Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "計測設定.ini")
    Call WriteIni("計測時間", "次回計測時間", Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "計測設定.ini")
    Call WriteIni("計測時間", "計測インターバル", toTMSstring(KE_intv), CurrentDir & "計測設定.ini")

    Call WriteIni("計測時間", "前回記録時間", Format$(kiroku_TimeZ, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "計測設定.ini")
    Call WriteIni("計測時間", "次回記録時間", Format$(kiroku_Time, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "計測設定.ini")
    Call WriteIni("計測時間", "記録インターバル", toTMSstring(kiroku_intv), CurrentDir & "計測設定.ini")
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

Public Function SEEKmoji(strCheckString As String, mojiST As Integer, mojiMAX As Integer) As String

    'Forカウンタ
    Dim i As Long
    '調査対象文字列の長さを格納
    Dim lngCheckSize As Long
    'ANSIへの変換後の文字を格納
    Dim lngANSIStr As Long
    
    Dim co As Integer '文字数
    Dim ss As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: ss = ""
    For i = 1 To lngCheckSize
        'StrConvでUnicodeからANSIへと変換
        lngANSIStr = LenB(StrConv(Mid$(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            ss = ss + Mid$(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = ss
End Function

Public Sub WriteEvents(st As String)
    Dim f As Integer
    f = FreeFile
    Open CurrentDir & App.EXEName & ".log" For Append Lock Write As #f
        'Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : " & st
        Print #f, st
    Close #f
    With MainForm.List1
        .AddItem st
        If 21 < .ListCount Then
            .RemoveItem (0)
        End If
        .ListIndex = .ListCount - 1
        .ListIndex = -1
    End With
End Sub

Public Sub WriteKanriLOG(keihou_L As Integer)
    Dim f As Integer
    
    f = FreeFile
    Open CurrentDir & "kanri.log" For Append Lock Write As #f
        Print #f, Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss") & " : " & StrConv(CStr(keihou_L), vbWide) & "次管理値を超えました。"
    Close #f

End Sub

Public Sub strTOtime(SS2 As String, tms As Date)
'時分秒文字列を分解して、日付型変数を作成

    Dim i As Integer
    Dim j As Integer
    Dim co As Integer
    Dim No As Integer
    Dim hh As Integer
    Dim mm As Integer
    Dim ss As Integer
    
    j = 0: co = 0
    Do
        i = InStr(j + 1, SS2, ":")
        co = co + 1
        If co = 1 Then
            If i = 0 Then
                No = CInt(Mid$(SS2, j + 1, Len(SS2) - j))
            Else
                hh = CInt(Mid$(SS2, j + 1, i - j - 1))
            End If
        Else
            If i = 0 Then
                If co = 2 Then
                    mm = CInt(Mid$(SS2, j + 1, Len(SS2) - j))
                Else
                    ss = CInt(Mid$(SS2, j + 1, Len(SS2) - j))
                End If
            Else
                If co = 2 Then
                    mm = CInt(Mid$(SS2, j + 1, i - j - 1))
                Else
                End If
            End If
        End If
        If i = 0 Then Exit Do
        j = i
    Loop
    
    Dim hd As Integer
    hd = hh \ 24
    hh = hh Mod 24
    tms = hd + TimeSerial(hh, mm, ss)
End Sub

Public Sub nextDate(da As Date, iv As Date, nd As Date)
'次回計測時間計算
'da:前回
'nd:次回
'iv:インターバル時間

    Dim hh0 As Integer, mm0 As Integer, ss0 As Integer
    Dim hh1 As Integer, mm1 As Integer, SS1 As Integer
    hh0 = Hour(da)
    mm0 = Minute(da)
    ss0 = Second(da)
    If iv < 1 Then
        hh1 = Hour(iv)
    Else
        hh1 = Fix(iv * 24)
    End If
    mm1 = Minute(iv)
    SS1 = Second(iv)
    Dim dds1
    
        dds1 = DateDiff("s", da, Now)
        dds1 = dds1 \ (hh1 * 60# * 60# + mm1 * 60# + SS1)
        dds1 = dds1 * (hh1 * 60# * 60# + mm1 * 60# + SS1)
        dds1 = DateAdd("s", dds1, da)
        dds1 = dds1 + iv 'KE_intv
    nd = CDate(dds1)
End Sub

Public Function toTMSstring(iv As Date) As String
'時刻文字列の作成

    Dim hh1 As Integer, mm1 As Integer, SS1 As Integer
    If iv < 1 Then
        hh1 = Hour(iv)
    Else
        hh1 = Fix(iv * 24)
    End If
    mm1 = Minute(iv)
    SS1 = Second(iv)
    
    Dim siv As String
    siv = Format(hh1, "") & ":" & Format$(mm1, "00") & ":" & Format(SS1, "00")
    If InStr(siv, ":") < 3 Then
        siv = "0" & Trim(siv)
    End If
    toTMSstring = siv
End Function

'深い階層のフォルダ作成
Public Function MakeFolder(strPath As String) As Boolean
'strPath = "c:\2013\09\30"
'などとしてコールすると、一度の最下層まで作成する
'
'
    On Error GoTo ErrHandler
    Dim pos1 As Integer
    Dim pos2 As Integer
    
    Dim temp As String
    Dim i As Integer
    Dim end_flg As Boolean
    
    end_flg = False
    For i = 1 To Len(strPath)
        pos1 = InStr(i, strPath, "\")
        pos2 = InStr(pos1 + 1, strPath, "\")
        If pos2 = 0 Then
            pos2 = Len(strPath)
            end_flg = True
        Else
            pos2 = pos2 - 1
        End If
        
        temp = Left$(strPath, pos2)
        If Dir$(temp, vbDirectory) = "" Then
            MkDir temp
        End If
        If end_flg = True Then
            Exit For
        Else
            i = pos1
        End If
    Next i
    MakeFolder = True
    Exit Function
ErrHandler:
    MakeFolder = False
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

Public Function Day2Dir(sd As String, ret As String) As Boolean
    'sd 日付 YYYY/MM/DD
    Dim YYY As Integer
    Dim mmm As Integer
    Dim ddd As Integer

    On Error GoTo Day2Dir9999

    YYY = Year(CDate(sd))
    mmm = Month(CDate(sd))
    ddd = Day(CDate(sd))
    ret = Format$(YYY, "0000")
    ret = ret & "\" & Format$(mmm, "00")
    ret = ret & "\" & Format$(ddd, "00")
    Day2Dir = True
    Exit Function
    
Day2Dir9999:
    On Error GoTo 0
    Day2Dir = False
End Function

'[ファイルの有無を調査する]
'
'■ 引数
'FileName:フルパスのファイル名
'□ 戻り値:パスの有無(True=存在する ,False=存在しない)
Public Function FileExists(ByVal FILENAME As String) As Boolean
  Dim TempAttr As Integer

  If (Len(FILENAME) = 0) Or (InStr(FILENAME, "*") > 0) Or _
                                                 (InStr(FILENAME, "?") > 0) Then
     FileExists = False
     Exit Function
  End If
  On Error GoTo ErrorFileExist
  ' ファイルの属性を得る
  TempAttr = GetAttr(FILENAME)
  ' ディレクトリであるかどうか調べる
  FileExists = ((TempAttr And vbDirectory) = 0)
  GoTo ExitFileExist
ErrorFileExist:
  FileExists = False
  Resume ExitFileExist
ExitFileExist:
  On Error GoTo 0
End Function

Public Sub WriteLog(st As String)
'st 説明文
    Dim f As Integer
    
    On Error GoTo WriteLog9999
    
    f = FreeFile
    Open App.Path & "\" & LOGFILE For Append As #f
    Print #f, Format$(Now, "YYYY/MM/DD hh:mm:ss"); " : ";
    Print #f, st
    Close #f

WriteLog9999:
    On Error GoTo 0
End Sub

