VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeiji1a 
   Caption         =   "Form3"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14385
   LinkTopic       =   "Form3"
   ScaleHeight     =   9015
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command5 
      Caption         =   "手動計測"
      Height          =   615
      Left            =   8040
      TabIndex        =   8
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame5 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin xTextNLib.xTextN xTextN2 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   2415
         _Version        =   262145
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         Locked          =   -1  'True
         PMenuCaption0   =   "元に戻す(&U)"
         PEnabled0       =   -1  'True
         PHidden0        =   0   'False
         PSeparator0     =   0   'False
         PMenuCaption1   =   "切り取り(&T)"
         PEnabled1       =   -1  'True
         PHidden1        =   0   'False
         PSeparator1     =   0   'False
         PMenuCaption2   =   "ｺﾋﾟｰ(&C)"
         PEnabled2       =   -1  'True
         PHidden2        =   0   'False
         PSeparator2     =   0   'False
         PMenuCaption3   =   "貼り付け(&P)"
         PEnabled3       =   -1  'True
         PHidden3        =   0   'False
         PSeparator3     =   0   'False
         PMenuCaption4   =   "削除(&D)"
         PEnabled4       =   -1  'True
         PHidden4        =   0   'False
         PSeparator4     =   0   'False
         PMenuCaption5   =   "すべて選択(&A)"
         PEnabled5       =   -1  'True
         PHidden5        =   0   'False
         PSeparator5     =   0   'False
         PMenuCaption6   =   ""
         PEnabled6       =   -1  'True
         PHidden6        =   -1  'True
         PSeparator6     =   0   'False
         PMenuCaption7   =   ""
         PEnabled7       =   -1  'True
         PHidden7        =   -1  'True
         PSeparator7     =   0   'False
         PMenuCaption8   =   ""
         PEnabled8       =   -1  'True
         PHidden8        =   -1  'True
         PSeparator8     =   0   'False
         PMenuCaption9   =   ""
         PEnabled9       =   -1  'True
         PHidden9        =   -1  'True
         PSeparator9     =   0   'False
         PMenuCaption10  =   ""
         PEnabled10      =   -1  'True
         PHidden10       =   -1  'True
         PSeparator10    =   0   'False
         PMenuCaption11  =   ""
         PEnabled11      =   -1  'True
         PHidden11       =   -1  'True
         PSeparator11    =   0   'False
         PMenuCaption12  =   ""
         PEnabled12      =   -1  'True
         PHidden12       =   -1  'True
         PSeparator12    =   0   'False
         PMenuCaption13  =   ""
         PEnabled13      =   -1  'True
         PHidden13       =   -1  'True
         PSeparator13    =   0   'False
         PMenuCaption14  =   ""
         PEnabled14      =   -1  'True
         PHidden14       =   -1  'True
         PSeparator14    =   0   'False
         PMenuCaption15  =   ""
         PEnabled15      =   -1  'True
         PHidden15       =   -1  'True
         PSeparator15    =   0   'False
         PMenuCaption16  =   ""
         PEnabled16      =   -1  'True
         PHidden16       =   -1  'True
         PSeparator16    =   0   'False
         PMenuCaption17  =   ""
         PEnabled17      =   -1  'True
         PHidden17       =   -1  'True
         PSeparator17    =   0   'False
         PMenuCaption18  =   ""
         PEnabled18      =   -1  'True
         PHidden18       =   -1  'True
         PSeparator18    =   0   'False
         PMenuCaption19  =   ""
         PEnabled19      =   -1  'True
         PHidden19       =   -1  'True
         PSeparator19    =   0   'False
      End
      Begin xTextNLib.xTextN xTextN2 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   2415
         _Version        =   262145
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   2
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         Locked          =   -1  'True
         PMenuCaption0   =   "元に戻す(&U)"
         PEnabled0       =   -1  'True
         PHidden0        =   0   'False
         PSeparator0     =   0   'False
         PMenuCaption1   =   "切り取り(&T)"
         PEnabled1       =   -1  'True
         PHidden1        =   0   'False
         PSeparator1     =   0   'False
         PMenuCaption2   =   "ｺﾋﾟｰ(&C)"
         PEnabled2       =   -1  'True
         PHidden2        =   0   'False
         PSeparator2     =   0   'False
         PMenuCaption3   =   "貼り付け(&P)"
         PEnabled3       =   -1  'True
         PHidden3        =   0   'False
         PSeparator3     =   0   'False
         PMenuCaption4   =   "削除(&D)"
         PEnabled4       =   -1  'True
         PHidden4        =   0   'False
         PSeparator4     =   0   'False
         PMenuCaption5   =   "すべて選択(&A)"
         PEnabled5       =   -1  'True
         PHidden5        =   0   'False
         PSeparator5     =   0   'False
         PMenuCaption6   =   ""
         PEnabled6       =   -1  'True
         PHidden6        =   -1  'True
         PSeparator6     =   0   'False
         PMenuCaption7   =   ""
         PEnabled7       =   -1  'True
         PHidden7        =   -1  'True
         PSeparator7     =   0   'False
         PMenuCaption8   =   ""
         PEnabled8       =   -1  'True
         PHidden8        =   -1  'True
         PSeparator8     =   0   'False
         PMenuCaption9   =   ""
         PEnabled9       =   -1  'True
         PHidden9        =   -1  'True
         PSeparator9     =   0   'False
         PMenuCaption10  =   ""
         PEnabled10      =   -1  'True
         PHidden10       =   -1  'True
         PSeparator10    =   0   'False
         PMenuCaption11  =   ""
         PEnabled11      =   -1  'True
         PHidden11       =   -1  'True
         PSeparator11    =   0   'False
         PMenuCaption12  =   ""
         PEnabled12      =   -1  'True
         PHidden12       =   -1  'True
         PSeparator12    =   0   'False
         PMenuCaption13  =   ""
         PEnabled13      =   -1  'True
         PHidden13       =   -1  'True
         PSeparator13    =   0   'False
         PMenuCaption14  =   ""
         PEnabled14      =   -1  'True
         PHidden14       =   -1  'True
         PSeparator14    =   0   'False
         PMenuCaption15  =   ""
         PEnabled15      =   -1  'True
         PHidden15       =   -1  'True
         PSeparator15    =   0   'False
         PMenuCaption16  =   ""
         PEnabled16      =   -1  'True
         PHidden16       =   -1  'True
         PSeparator16    =   0   'False
         PMenuCaption17  =   ""
         PEnabled17      =   -1  'True
         PHidden17       =   -1  'True
         PSeparator17    =   0   'False
         PMenuCaption18  =   ""
         PEnabled18      =   -1  'True
         PHidden18       =   -1  'True
         PSeparator18    =   0   'False
         PMenuCaption19  =   ""
         PEnabled19      =   -1  'True
         PHidden19       =   -1  'True
         PSeparator19    =   0   'False
      End
      Begin xTextNLib.xTextN xTextN2 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   2415
         _Version        =   262145
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         Locked          =   -1  'True
         PMenuCaption0   =   "元に戻す(&U)"
         PEnabled0       =   -1  'True
         PHidden0        =   0   'False
         PSeparator0     =   0   'False
         PMenuCaption1   =   "切り取り(&T)"
         PEnabled1       =   -1  'True
         PHidden1        =   0   'False
         PSeparator1     =   0   'False
         PMenuCaption2   =   "ｺﾋﾟｰ(&C)"
         PEnabled2       =   -1  'True
         PHidden2        =   0   'False
         PSeparator2     =   0   'False
         PMenuCaption3   =   "貼り付け(&P)"
         PEnabled3       =   -1  'True
         PHidden3        =   0   'False
         PSeparator3     =   0   'False
         PMenuCaption4   =   "削除(&D)"
         PEnabled4       =   -1  'True
         PHidden4        =   0   'False
         PSeparator4     =   0   'False
         PMenuCaption5   =   "すべて選択(&A)"
         PEnabled5       =   -1  'True
         PHidden5        =   0   'False
         PSeparator5     =   0   'False
         PMenuCaption6   =   ""
         PEnabled6       =   -1  'True
         PHidden6        =   -1  'True
         PSeparator6     =   0   'False
         PMenuCaption7   =   ""
         PEnabled7       =   -1  'True
         PHidden7        =   -1  'True
         PSeparator7     =   0   'False
         PMenuCaption8   =   ""
         PEnabled8       =   -1  'True
         PHidden8        =   -1  'True
         PSeparator8     =   0   'False
         PMenuCaption9   =   ""
         PEnabled9       =   -1  'True
         PHidden9        =   -1  'True
         PSeparator9     =   0   'False
         PMenuCaption10  =   ""
         PEnabled10      =   -1  'True
         PHidden10       =   -1  'True
         PSeparator10    =   0   'False
         PMenuCaption11  =   ""
         PEnabled11      =   -1  'True
         PHidden11       =   -1  'True
         PSeparator11    =   0   'False
         PMenuCaption12  =   ""
         PEnabled12      =   -1  'True
         PHidden12       =   -1  'True
         PSeparator12    =   0   'False
         PMenuCaption13  =   ""
         PEnabled13      =   -1  'True
         PHidden13       =   -1  'True
         PSeparator13    =   0   'False
         PMenuCaption14  =   ""
         PEnabled14      =   -1  'True
         PHidden14       =   -1  'True
         PSeparator14    =   0   'False
         PMenuCaption15  =   ""
         PEnabled15      =   -1  'True
         PHidden15       =   -1  'True
         PSeparator15    =   0   'False
         PMenuCaption16  =   ""
         PEnabled16      =   -1  'True
         PHidden16       =   -1  'True
         PSeparator16    =   0   'False
         PMenuCaption17  =   ""
         PEnabled17      =   -1  'True
         PHidden17       =   -1  'True
         PSeparator17    =   0   'False
         PMenuCaption18  =   ""
         PEnabled18      =   -1  'True
         PHidden18       =   -1  'True
         PSeparator18    =   0   'False
         PMenuCaption19  =   ""
         PEnabled19      =   -1  'True
         PHidden19       =   -1  'True
         PSeparator19    =   0   'False
      End
      Begin VB.Label Label6 
         Caption         =   "次回測定時間"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "今回測定時間"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "計測値"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   810
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   8715
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18812
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5997
            MinWidth        =   5997
            Text            =   "2013/07/10"
            TextSave        =   "2013/07/10"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmKeiji1a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim STtime As Date, minTIME As Date
Dim keisoku_f As Boolean
    Dim Thistime As String
    Dim stat As Integer


Private Sub Timer1_Timer()
    Call tSub
End Sub

Private Sub Command5_Click()
    Dim i As Integer, f As Integer
    Dim ckTIME As Date
    Dim stat As Integer
    Dim Thistime As String
    Dim Syudou As Boolean
    Dim Z_jidou_time As Date
    Dim minTIME As Date
    Dim t1 As Date, t2 As Date
    
    On Error GoTo KeiERR

    ckTIME = DateAdd("n", -20, Keisoku_Time)
    If DateDiff("s", Now, ckTIME) <= 0 Then
        MsgBox "次回測定時間が終了するまで、手動計測はできません。", vbCritical, "警告メッセージ"
        Exit Sub
    End If
    
    If vbOK = MsgBox("「OK」をクリックすると、計測します｡", vbOKCancel + vbInformation, "手動計測") Then
        Syudou = True
    Else
        Syudou = False
    End If
    If Syudou = True Then
        StatusBar1.Panels(1).Text = "*** 手動計測中 ***"
        Enabled = False
        
        Form2.Label1.Caption = "*** 手 動 計 測 中 ***"
        Form2.Show
        
        MDY = Now
        Thistime = Format$(MDY, "yyyy/mm/dd hh:nn:ss")
        
        If Command$ = "TEST" Then
            RsctlFrm.StatusBar1.Panels(1).Text = "スイッチＯＮ"
            Call DAMMYdate(stat, Thistime, 0)
        Else
            'GTS-8のスイッチON
            RsctlFrm.StatusBar1.Panels(1).Text = "スイッチＯＮ"
            Call frmGTS800A.GTS8on
            
            RsctlFrm.StatusBar1.Panels(1).Text = "初期設定"
            Call frmGTS800A.GTS8init 'GTS-8の初期設定
            '
            Call frmGTS800A.SOKUTEI(KEISOKU.Data_path & DATA_DAT)
            
            'GTS-8のスイッチOFF
            RsctlFrm.StatusBar1.Panels(1).Text = "スイッチＯＦＦ"
            Call frmGTS800A.GTS8off
        End If
        Unload Form2
        
''        Z_jidou_time = Z_Keisoku_Time
        Z_Keisoku_Time = CDate(Thistime)
        
        keihou_L = 0
        Call DataPrint
        
        '警報発令
        If keihou_L > 1 Then
            f = FreeFile
            Open CurrentDir & "kanri.log" For Append Lock Write As #f
                Print #f, Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss") & " : " & StrConv(CStr(keihou_L), vbWide) & "次管理値を超えました。"
            Close #f
        End If
        
        If DateDiff("s", Kkeiji.ED, Z_Keisoku_Time) > 0 Then
            Call KeijiInit
        Else
            Call KeijiPlot2
        End If
        
        '次回計測時間計算
        MDY = Now
        KE_intv = CDate(Lebel_intv(keihou_L) & ":00") 'インターバル時間
        minTIME = DateAdd("m", 1, MDY)
        For i = 1 To 24 / Lebel_intv(keihou_L)
            t1 = DateValue(MDY) + Lebel_time(keihou_L, i)
            If DateDiff("s", t1, MDY) > 0 Then t1 = DateAdd("d", 1, t1)
            If DateDiff("s", t1, minTIME) > 0 Then minTIME = t1
        Next i
        Keisoku_Time = minTIME                        '次回計測時間
        Call IntvWrite
        
        xTextN2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
        xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
        
        Enabled = True
        If stat = 0 Then StatusBar1.Panels(1).Text = ""
''        Z_Keisoku_Time = Z_jidou_time
    End If
    
    Exit Sub

KeiERR:
    Close
    Unload Form2
'''    Unload Form1
    
    If Err.Number = 10000 Then
        MsgBox "通信エラー"
    Else
        MsgBox "エラー:" & Err.Number
    End If

End Sub

Private Sub Form_Load()
'   Me.Height = 15360 - 420  '11000 '16590
'   Me.Width = 19200 '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
'    Me.Height = 12000 '16590
'    Me.Width = 16000 '21210
'    Left = (Screen.Width - Me.Width) / 2
'    Top = 0

    Dim rc As Integer
    Dim t1 As Date, t2 As Date
    Dim i As Integer

    '2001/05/16
    t1 = Timer
    frmInitMsg.Show: frmInitMsg.Refresh
    If Command$ = "TEST" Then
        Do
            DoEvents
            t2 = Timer
            If (t2 - t1) > 2 Then Exit Do
        Loop
    Else
        If frmGTS800A.GTS8on = 0 Then
            MsgBox "GTS-800の反応がありません。" & CStr(rc), vbCritical
            Close
            End
        End If
        Call frmGTS800A.GTS8init 'GTS-8の初期設定
    End If
    Unload frmInitMsg

    '自動計測画面表示
    'Show
'    If Command$ = "CHECK" Then
'        Command3.Visible = True
'    Else
'        Command3.Visible = False
'    End If
'    If Command$ = "" Then
'        Command4.Visible = False
'    Else
'        Command4.Visible = True
'    End If
    'Refresh
        
    Me.Enabled = False
'    Call HyoujiInit
    keihou_L = 0
    Call DataPrint
    Me.Enabled = True
    
    '次回計測時間計算
    KE_intv = CDate(Lebel_intv(keihou_L) & ":00") 'インターバル時間
    minTIME = DateAdd("m", 1, Now)
    For i = 1 To 24 / Lebel_intv(keihou_L)
        t1 = DateValue(Now) + Lebel_time(keihou_L, i)
        If DateDiff("s", t1, Now) > 0 Then t1 = DateAdd("d", 1, t1)
        If DateDiff("s", t1, minTIME) > 0 Then minTIME = t1
    Next i
    Keisoku_Time = minTIME                        '次回計測時間
    Call IntvWrite
    xTextN2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
    '経時変化図画面表示
    Kkeiji.Yti = GetIni("常時表示設定", "Ｙ軸タイトル", CurrentDir & "計測設定.ini")
    Kkeiji.Yu = GetIni("常時表示設定", "Ｙ軸単位", CurrentDir & "計測設定.ini")
    
    Kkeiji.Xmax = GetIni("常時表示設定", "Ｘ軸日数", CurrentDir & "計測設定.ini")
    Kkeiji.xBUN = GetIni("常時表示設定", "Ｘ軸分割数", CurrentDir & "計測設定.ini")
    If Kkeiji.Xmax = 0 Then Kkeiji.Xmax = 4: Call WriteIni("常時表示設定", "Ｘ軸日数", CStr(Kkeiji.Xmax), CurrentDir & "計測設定.ini")
    If Kkeiji.xBUN = 0 Then Kkeiji.xBUN = 12: Call WriteIni("常時表示設定", "Ｘ軸分割数", CStr(Kkeiji.xBUN), CurrentDir & "計測設定.ini")
    
    Kkeiji.YMIN = CSng(GetIni("常時表示設定", "Ｙ軸最小値", CurrentDir & "計測設定.ini"))
    Kkeiji.YMAX = CSng(GetIni("常時表示設定", "Ｙ軸最大値", CurrentDir & "計測設定.ini"))
    Kkeiji.yBUN = CSng(GetIni("常時表示設定", "Ｙ軸分割数", CurrentDir & "計測設定.ini"))
    If Kkeiji.yBUN = 0 Then
        Kkeiji.yBUN = 10
        Call WriteIni("常時表示設定", "Ｙ軸分割数", CStr(Kkeiji.yBUN), CurrentDir & "計測設定.ini")
    End If
    If Kkeiji.YMIN = Kkeiji.YMAX Then
        Kkeiji.YMIN = 0
        Kkeiji.YMAX = 400
        Call WriteIni("常時表示設定", "Ｙ軸最小値", CStr(Kkeiji.YMIN), CurrentDir & "計測設定.ini")
        Call WriteIni("常時表示設定", "Ｙ軸最大値", CStr(Kkeiji.YMAX), CurrentDir & "計測設定.ini")
    End If
    Call KeijiInit
    
    '自動計測開始
    keisoku_f = False
'    Me.SetFocus
    Timer1.Interval = 200
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Ret As Long
    Dim RetString As String
    Dim i As Integer, ENDsw As Boolean, f As Integer
    Dim rc As Integer
    
    On Error Resume Next
    
    If UnloadMode <= 1 Then
        If vbCancel = MsgBox("「OK」をクリックすると、計測が終了します｡", vbOKCancel + vbExclamation, "終了の確認") Then
            Cancel = True
            ENDsw = False
        Else
            ENDsw = True
        End If
    Else
        ENDsw = True
    End If
    
    If ENDsw = True Then
        
        '終了ログ
        f = FreeFile
        Open CurrentDir & "PRG-event.log" For Append Lock Write As #f
            Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 終了"
        Close #f
        
        Call IntvWrite
        
        Close
        
        While Forms.Count > 1
            '---自分以外のフォームを探します
            i = 0
            While Forms(i) Is Me
                i = i + 1
            Wend
            Unload Forms(i)
        Wend
        
        '---自分自身もアンロードし、アプリケーションは終了します
        Unload Me
        End
    End If

End Sub



Private Sub tSub()
    Dim i As Integer
    Dim t1 As Date
    
        MDY = Now
        Thistime = Format$(MDY, "yyyy/mm/dd hh:nn:ss")
        MainForm.StatusBar1.Panels(2).Text = Format$(Thistime, "yyyy年mm月dd日 hh時nn分ss秒")
        
        If Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss") = Thistime Then keisoku_f = True
        
        If keisoku_f = True Then
            MainForm.StatusBar1.Panels(1).Text = "*** 計測中 ***"
            MainForm.Enabled = False
                
            '    Tensu% = 3
            '    Tensu% = InitDT.co
            '    HeikinKaisuu = InitDT.AvgCO '= 3
            '    x0 = InitDT.Kx
            '    y0 = InitDT.Ky
            '    z0 = InitDT.Kz
            '    MH = InitDT.Kmh
            '    x1 = InitDT.Bx
            '    y1 = InitDT.By
            '    z1 = InitDT.Bz
            '    For i = 1 To InitDT.co
            '        H#(1, i) = PoDT(i).Hdt
            '        V#(1, i) = PoDT(i).Vdt
            '        S#(1, i) = PoDT(i).Sdt
            '    Next
            '    AZIMUTH = InitDT.HOKO
            Form2.Show
            If Command$ = "TEST" Then
                Call DAMMYdate(stat, Thistime, 0)
            Else
                'GTS-8のスイッチON
                If frmGTS800A.GTS8on = 0 Then
                    GoTo 99
                End If
                Call frmGTS800A.GTS8init 'GTS-8の初期設定
                '
                Call frmGTS800A.SOKUTEI(KEISOKU.Data_path & DATA_DAT)
                
                'GTS-8のスイッチOFF
                Call frmGTS800A.GTS8off
            End If
            Unload Form2
            
            keihou_L = 0
            Call DataPrint
            
            '次回計測時間計算
            If keisoku_f = True Then
                keisoku_f = False
                Z_Keisoku_Time = Keisoku_Time
                
                'Keisoku_Time = CDate(Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")) + KE_intv
                KE_intv = CDate(Lebel_intv(keihou_L) & ":00") 'インターバル時間
                minTIME = DateAdd("m", 1, MDY)
                For i = 1 To 24 / Lebel_intv(keihou_L)
                    t1 = DateValue(MDY) + Lebel_time(keihou_L, i)
                    If DateDiff("s", t1, MDY) > 0 Then t1 = DateAdd("d", 1, t1)
                    If DateDiff("s", t1, minTIME) > 0 Then minTIME = t1
                Next i
                Keisoku_Time = minTIME                        '次回計測時間
            End If
            Call IntvWrite


            '警報発令
            If keihou_L > 1 Then
                Call WriteKanriLOG(keihou_L)
            End If
            
            If DateDiff("s", Kkeiji.ED, Z_Keisoku_Time) > 0 Then
                Call KeijiInit
            Else
                Call KeijiPlot2
            End If
            
'            MainForm.xTextN2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'            MainForm.xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            
            MainForm.Enabled = True
            If stat = 0 Then MainForm.StatusBar1.Panels(1).Text = ""
        End If
            
        '管理レベル点滅
'        If keihou_L = 1 Then
'            MainForm.Shape1.BackColor = &H80FF80
'        Else
'            If Second(CDate(Thistime)) Mod 2 = 0 Then
'                MainForm.Shape1.BackColor = QBColor(15)
'            Else
'                If keihou_L = 2 Then MainForm.Shape1.BackColor = &H80FFFF
'                If keihou_L = 3 Then MainForm.Shape1.BackColor = &HFF80FF
'                If keihou_L = 4 Then MainForm.Shape1.BackColor = RGB(256, 60, 60)
'
'            End If
'        End If
        
        
99      '計測時間がすぎた場合
        If DateDiff("s", Keisoku_Time, MDY) >= 0 Then  'If nt < Now Then
            'Keisoku_Time = T_ajt(Z_Keisoku_Time, KE_intv)
            
            KE_intv = CDate(Lebel_intv(keihou_L) & ":00") 'インターバル時間
            minTIME = DateAdd("m", 1, MDY)
            For i = 1 To 24 / Lebel_intv(keihou_L)
                t1 = DateValue(MDY) + Lebel_time(keihou_L, i)
                If DateDiff("s", t1, MDY) > 0 Then t1 = DateAdd("d", 1, t1)
                If DateDiff("s", t1, minTIME) > 0 Then minTIME = t1
            Next i
            Keisoku_Time = minTIME                        '次回計測時間
            Call IntvWrite
            
            xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            
            'ログ
            Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 計測時間が過ぎていたため、再設定しました。")
        End If

End Sub

