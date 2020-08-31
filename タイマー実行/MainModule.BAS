Attribute VB_Name = "MainModule"
Option Explicit

'---------------------------------------------------------------
'自動計測
'---------------------------------------------------------------
'作成履歴
'---------------------------------------------------------------
'修正履歴
'---------------------------------------------------------------
Public Const PI As Double = 3.1415926535898

'定数

'計測インターバル設定
Public Keisoku_Time As Date        '次回計測
Public KE_intv As Date             '計測インターバル
Public Keisoku_TimeZ As Date      '前回計測
Public Lebel_time(4, 24) As Date   'レベル別計測時間
Public Lebel_intv(4) As Integer    'レベル別計測インターバル

'フォームサイズ変更時のコマンドボタンのフォームの枠からの位置
Public Const mintControlMargin = 60

Public RsctlFrm As Object

Public ivCode '0:計測、1:記録

Public MDY As Date      '計測日時

Public CurrentDir As String

'---------------------------------------------------------------------
Public Sub Main()
    Dim f As Integer, L As String
    Dim i As Integer, t_ID As Integer, k_ID As Integer
    Dim j As Integer, co As Integer
    Dim kiroku_f As Boolean
    Dim ret As Integer, RetString As String
    Dim SS1 As String, SS2 As String
    Dim t1 As Date, t2 As Date
    Dim rc As Integer
    Dim STtime As Date, minTIME As Date
    

'Debug.Print GetPathNameToFullPas("C:\Windows\notepad.exe")
'Debug.Print GetFullPasToFileName("C:\Windows\notepad.exe")


'    If App.PrevInstance = True Then
'        MsgBox "既に起動しています。", vbCritical, "起動エラー"
'        End
'    End If
    
  '  On Error GoTo TRAP_Main
    
'    ChDrive App.Path
'    ChDir App.Path
    SetCurrentDirectory App.Path

    
    CurrentDir = App.Path
    If Right(CurrentDir, 1) = "\" Then Else CurrentDir = CurrentDir & "\"
    
    '起動ログ
    f = FreeFile
    Open CurrentDir & "PRG-event.log" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 起動"
    Close #f
    
'    'ファイル読み込み
'    KEISOKU.Data_path = GetIni("フォルダ名", "計測データ", CurrentDir & "ExecTimer.ini")
'    KEISOKU.Tabl_path = GetIni("フォルダ名", "環境データ", CurrentDir & "ExecTimer.ini")
'    KEISOKU.BackupPath = GetIni("フォルダ名", "Backup", CurrentDir & "ExecTimer.ini")
'    If Right$(KEISOKU.BackupPath, 1) <> "\" Then
'        KEISOKU.BackupPath = KEISOKU.BackupPath & "\"
'    End If
    
    Dim sKE_intv As String
    Keisoku_TimeZ = CDate(GetIni("計測時間", "前回計測時間", CurrentDir & "ExecTimer.ini"))
    Keisoku_Time = CDate(GetIni("計測時間", "次回計測時間", CurrentDir & "ExecTimer.ini"))
    sKE_intv = (GetIni("計測時間", "計測インターバル", CurrentDir & "ExecTimer.ini"))
    Call strTOtime(sKE_intv, KE_intv)
    If KE_intv = 0 Then
        KE_intv = #1:00:00 AM#
        Call WriteIni("計測時間", "計測インターバル", Format$(KE_intv, "hh:nn:ss"), CurrentDir & "ExecTimer.ini")
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
   Dim tmp As String
   
   Span = Num \ 2
   Do While Span > 0
      For i = Span To Num - 1
         j% = i% - Span + 1
         For j = (i - Span + 1) To 1 Step -Span
            If Aray(j) >= Aray(j + Span) Then Exit For
            ' 順番の異なる配列要素を入れ替えます.
            tmp = Aray(j)
            Aray(j) = Aray(j + Span)
            Aray(j + Span) = tmp
         Next j
      Next i
      Span = Span \ 2
   Loop
End Sub

Public Sub IntvWrite()
    Call WriteIni("計測時間", "前回計測時間", Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "ExecTimer.ini")
    Call WriteIni("計測時間", "次回計測時間", Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss"), CurrentDir & "ExecTimer.ini")
    Call WriteIni("計測時間", "計測インターバル", toTMSstring(KE_intv), CurrentDir & "ExecTimer.ini")

End Sub

Public Function FileCheck(FileName As String, FileTitle As String) As Integer
    Dim i As Integer

    On Error Resume Next

    i = 0
    If Dir$(FileName) = "" Then Else i = 1
    If i = 0 Then
        MsgBox FileTitle & "ファイル(" & FileName & ")が見つかりません。確認してください。", vbCritical, "エラーメッセージ"
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

Public Sub KEISAN(pd() As Double, Kpd As Double)
    Dim AnsSNG As Single, AnsDBL As Double
    Dim ten_ID As Integer
    
''    If pd(3) >= 999999 Or pd(6) >= 999999 Then
''        AnsDBL = 999999
''    Else
''        AnsDBL = pd(3) + 0.3 * pd(6)
''    End If
    If pd(3) >= 999999 Or pd(5) >= 999999 Then
        AnsDBL = 999999
    Else
        AnsDBL = pd(3) + 0.3 * pd(5)
    End If
    Kpd = AnsDBL
End Sub


Sub WriteEvents(st As String)
    Dim f As Integer
    f = FreeFile
    Open CurrentDir & "PRG-event.log" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : " & st
    Close #f
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
                No = CInt(Mid(SS2, j + 1, Len(SS2) - j))
            Else
                hh = CInt(Mid(SS2, j + 1, i - j - 1))
            End If
        Else
            If i = 0 Then
                If co = 2 Then
                    mm = CInt(Mid(SS2, j + 1, Len(SS2) - j))
                Else
                    ss = CInt(Mid(SS2, j + 1, Len(SS2) - j))
                End If
            Else
                If co = 2 Then
                    mm = CInt(Mid(SS2, j + 1, i - j - 1))
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
    siv = Format(hh1, "") & ":" & Format(mm1, "00") & ":" & Format(SS1, "00")
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

    temp = Left(strPath, pos2)
    If Dir(temp, vbDirectory) = "" Then
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

