VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Object = "{323DBF23-9372-4ADC-80FF-0ABA14A5F694}#4.2#0"; "xCBtnN.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "2"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9960
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   9960
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6000
      Top             =   240
   End
   Begin VB.Frame Frame4 
      Caption         =   "掘削位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1335
      Left            =   7560
      TabIndex        =   2
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Ａ断面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Ｂ断面"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   870
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   465
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   17515
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin xCBtnNLib.xCmdBtnN xCmdBtn1 
      Height          =   1095
      Left            =   14400
      TabIndex        =   1
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   262146
      _ExtentX        =   2566
      _ExtentY        =   1931
      _StockProps     =   77
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionStd      =   "警報停止"
      Picture         =   "MainForm.frx":0442
      ForeColor       =   255
      DownPicture     =   "MainForm.frx":045E
      DisabledPicture =   "MainForm.frx":047A
      OnPicture       =   "MainForm.frx":0496
      PMenuCaption0   =   "元に戻す(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   -1  'True
      PSeparator0     =   0   'False
      PMenuCaption1   =   "切り取り(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   -1  'True
      PSeparator1     =   0   'False
      PMenuCaption2   =   "ｺﾋﾟｰ(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   -1  'True
      PSeparator2     =   0   'False
      PMenuCaption3   =   "貼り付け(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   -1  'True
      PSeparator3     =   0   'False
      PMenuCaption4   =   "削除(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   -1  'True
      PSeparator4     =   0   'False
      PMenuCaption5   =   "すべて選択(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   -1  'True
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
   Begin VB.Label Label3 
      Caption         =   "yyyy/mm/dd hh:mm:ss"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "現在日時"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   25000
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   25000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu file 
      Caption         =   "ファイル"
      Begin VB.Menu end 
         Caption         =   "終了"
      End
   End
   Begin VB.Menu mnuCrt0 
      Caption         =   "表示"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuIntv 
      Caption         =   "ｲﾝﾀｰﾊﾞﾙ"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSet 
      Caption         =   "設定"
      Visible         =   0   'False
      Begin VB.Menu mnuTable 
         Caption         =   "環境設定"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim bTelDaial As Boolean  ' ダイアルアップ開始したらTrue
Dim iDaialID As Integer

Dim keihouF As Boolean    '警報ファイルがあったら True
Dim SendDataF As Boolean  '送信ファイルがあったら True
Dim SoushinFileCK As Boolean '送信用ファイルのチェックが済んだら　True
Dim SoushinCount As Integer

Dim RASenable As Boolean  '回線接続中 True
Dim SendComp As Boolean   '送信終了したら True


Private Sub Command1_Click()
    'TEST用
'            If CheckDataFile(TdsDataPath) <> 0 Then
'                iDaialID = 1
'                Call Soushin(iDaialID)
'            End If
    'keihouCK
End Sub

Private Sub end_Click()
    Unload Me 'End
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
      Unload Me
    End If
    
    Dim tp&, lt&
    On Local Error GoTo E01
    tp = CLng(GetIni("Form", "Top", mINIfile))
    If tp < 0 Then tp = 0
    lt = CLng(GetIni("Form", "left", mINIfile))
    If lt < 0 Then lt = 0
    
S01:
    Top = tp
    Left = lt
    On Local Error GoTo 0
    
    Me.Height = 1545 '2865
    Me.Width = 4545
    
    Caption = GetIni("現場名", "現場名", mINIfile) + " データ送信"
    
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    Timer1.Enabled = True
    
'    NoRAS = GetIni("RAS", "NoRAS", mINIfile)
'    Dim i As Integer
'    For i = 1 To 1
'        RAS(i).eName = GetIni("RAS", "eName" & i, mINIfile)
'        RAS(i).User = GetIni("RAS", "User" & i, mINIfile)
'        RAS(i).Pw = GetIni("RAS", "Pw" & i, mINIfile)
'        Call RasInitial(i)
'    Next i
    SoushinCount = 5
Exit Sub

E01:
    On Local Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rc%
    Dim i As Integer, ENDsw As Boolean, f As Integer
    Dim RetString As String
    
    On Error Resume Next
    
    If UnloadMode < 1 Then
        If vbCancel = MsgBox("「OK」をクリックすると、終了します｡", vbOKCancel + vbExclamation, "終了の確認") Then
            Cancel = True
            ENDsw = False
        Else
            ENDsw = True
        End If
    Else
        ENDsw = True
    End If
    
    If ENDsw = True Then
        
        Call WriteIni("form", "top", CStr(Top), mINIfile)
        Call WriteIni("form", "left", CStr(Left), mINIfile)
        
        '終了ログ
        f = FreeFile
        Open CurrentDir & "PRG-event.log" For Append Lock Write As #f
            Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 終了"
        Close #f
        
'        Call IntvWrite
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

Private Sub RasClient1_Connected()
'    Timer1.Enabled = False
'
'    '接続完了
'
'    Dim ret%
'
'    RASenable = True
'
'    If keihouF = True Then
'        SoushinCount = SoushinCount - 1
'        If 0 < SoushinCount Then
'            Call FileJoinSend
'            Call SendLogMove
'        Else
'            SoushinCount = 0
'        End If
'        keihouF = False
'    Else
'        SoushinCount = 5
'    End If
'
''    If SendDataF = True Then
''        Call FindDataFile(2, TdsDataPath(1))
''        Call SendPNG(ret)
''        SendDataF = False
''    End If
'
'    SendComp = True
'
'    RasClient1.HangUp  '電話を切る
'    Call Sleep(2000)
'    StatusBar1.Panels(1).Text = ""
'    SoushinFileCK = False
'    Timer1.Enabled = True
Exit Sub
            'Shape1.FillColor = RGB(0, &HFF, 0)
'            StatusBar1.Panels(1).Text = "接続中"
            'Call SendFTP(ret)
'            If RasClient1.Active = True Then
'                RasClient1.HangUp  '電話を切る
'                Call Sleep(2000)
'                ConnectCK = 0
'                StatusBar1.Panels(1).Text = ""
'            End If
'            If ret = -1 Then
'                'Call DelFile("\FTP\" & Trim(henkanTBL(1).FileName) & ".dat")
'                Call AllFileDelete
'            End If
'            Z_Keisoku_Time = Keisoku_Time
'            Keisoku_Time = CDate(Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")) + KE_intv
'            If rSettei = True Then
'                Keisoku_Time = CDate(GetIni("計測時間", "次回計測時間", MyAppPath & "settei.ini"))
'                KE_intv = CDate(GetIni("計測時間", "計測インターバル", MyAppPath & "settei.ini"))
'                Call IntvWrite
'                Text1(1) = Format$(KE_intv, "           hh:nn:ss")
'                Text1(2) = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'                Call Ktime_ck: Text1(2) = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'                KEISOKU.Waittime = CInt(GetIni("計測", "待ち", MyAppPath & "settei.ini"))
'                Call WriteIni("計測", "待ち", CStr(KEISOKU.Waittime), MyAppPath & "mSoushin.ini")
'                Call DelFile(App.Path & "\settei.ini")
'                rSettei = False
'            End If
'    Call FTPrw
End Sub

Private Sub exMail()
    Timer1.Enabled = False
    Dim ret%
    
    RASenable = True
    
    If keihouF = True Then
        Call WriteLog("警報メール送信開始")
        '2012/05/11 SoushinCount を無効に
        'SoushinCount = SoushinCount - 1
        'If 0 < SoushinCount Then
            Call FileJoinSend
            Call SendLogMove
        'Else
        '    SoushinCount = 0
        'End If
        keihouF = False
    Else
        SoushinCount = 5
    End If
    
    SendComp = True
    
'    RasClient1.HangUp  '電話を切る
'    Call Sleep(2000)
    StatusBar1.Panels(1).Text = ""
    SoushinFileCK = False
    Timer1.Enabled = True

End Sub

'Private Sub RasClient1_StatusChange(StatusMsg As String, ByVal StatusCode As Long)
'    StatusBar1.Panels(1).Text = str$(StatusCode) & " " & StatusMsg
'End Sub

'Private Sub RasClient1_StatusError(ErrorMsg As String, ByVal ErrorCode As Long)
'    'Form2.Label1 = Str$(ErrorCode) & ErrorMsg
'    'Form2.CmdCancel.Caption = "OK"
'    Dim f As Integer
'
'    ConnectCK = -3
'    If ErrorCode = 676 Then ConnectCK = -2    'ﾋﾞｼﾞｰ
'    If ErrorCode = 678 Then ConnectCK = -6    '応答がありません
'    If ErrorCode = 633 Then ConnectCK = -7    'ﾎﾟｰﾄは既に使用中か、ﾘﾓｰﾄ ｱｸｾｽ ﾀﾞｲﾔﾙｱｳﾄに対して構成されていません。
'
'    StatusBar1.Panels(1).Text = str$(ErrorCode) & " " & ErrorMsg
'
'    f = FreeFile
'    Open CurrentDIR & "modem-err.dat" For Append As #f
'        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " :";
'        Print #f, str$(ErrorCode) & " " & ErrorMsg
'    Close #f
'End Sub

'Private Function RasStart(ByVal id As Integer) As Boolean
''    Dim Ent As Entry
''    Set Ent = RasClient1.Items.Item(iEntNo)
''    tx5 = Ent.UserName
''    tx6 = Ent.Password
''    Set Ent = Nothing
'
'    'NTの場合Rasmon.exeを起動します。
'    On Error Resume Next
'    If m_Rasmon = 0 Then
'        m_Rasmon = Shell("rasmon.exe", vbNormalFocus)
'    End If
'    On Error GoTo 0
'
'    With MainForm.RasClient1
'        .ItemIndex = RAS(id).iEntNo
'        'プロパティーをセットします。
'        .ReDialTimes = 1
'        .ReDialInterval = 10
'        ''Form2.Caption = RasClient1.EntryName
'        .UserName = "a545352322@p.auone-net.jp" 'RAS(id).User '"ユーザー名を設定"
'        .Password = "yo2803ks" 'RAS(id).Pw '"パスワードを設定"
'    ''    If Check1.Value = 0 Then
'    ''        .UseSpValue = True
'    ''        .SpTelephoneNumber = Text2
'    ''        .SpDomainName = Text3
'    ''        .SpCallBackNumber = Text4
'    ''    Else
'            'ダイアルアップネットワークにあらかじめセットされた値が使用される
'            .UseSpValue = False
'    ''    End If
'
'        '接続します。
'        If .Connect(Me.hWnd) = -1 Then
'            '接続処理開始
'            RasStart = True
'        Else
'            ''接続処理に失敗
'            RasStart = False
'        End If
'    End With
'End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Dim Tstime As String
    Dim keisoku_f As Boolean
    Dim ret As Integer
    
    Tstime = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    
    If SoushinFileCK = False Then Call tmSub
'    If RasClient1.Active = True Then
'        RasClient1.HangUp  '電話を切る
'        Call Sleep(3000)
'    End If
'    If RASenable = True And SendComp = True Then
'        RasClient1.HangUp  '電話を切る
'        Call Sleep(2000)
'        RASenable = False
'        SendComp = False
'    End If
    
    Timer1.Enabled = True
End Sub

Private Sub tmSub()
    
    '警報データファイルがあったら、そのデータを送信
    'keihouF = keihouCK
        keihouF = FileExists(App.Path & "\" & KeihouFile)
    
'    '計測データフォルダの中身があったら、そのデータを送信
'    If CheckDataFile(TdsDataPath(1)) <> 0 Then
'        StatusBar1.Panels(1).Text = "送信開始-1"
'        waits 5
'        SendDataF = True
'
'        If SendDataF = True Then
'            Call FindDataFile(2, TdsDataPath(1), 1)
'            SendDataF = False
'        End If
'        StatusBar1.Panels(1).Text = ""
'
'    End If
'    If CheckDataFile(TdsDataPath(2)) <> 0 Then
'        StatusBar1.Panels(1).Text = "送信開始-2"
'        waits 5
'        SendDataF = True
'
'        If SendDataF = True Then
'            Call FindDataFile(2, TdsDataPath(2), 2)
'            SendDataF = False
'        End If
'        StatusBar1.Panels(1).Text = ""
'
'    End If

'    If keihouF = True Or SendDataF = True Then
    If keihouF = True Then
        '電話をかける
'        If NoRAS = 0 Then
'            If RasStart(2) = False Then Exit Sub
'            SoushinFileCK = True
'        Else
            Call exMail
'        End If
    End If

    Unload Me
    
Exit Sub

End Sub

'Private Sub Soushin(ByVal id As Integer)
'    Dim ModemSW As Boolean
'    Dim t1 As Date, t2 As Date
'
'    ConnectCK = 0
'    ModemSW = RasStart(id) '電話する
'    If ModemSW = False Then strData = "": MainForm.StatusBar1.Panels(1).Text = "": Exit Sub
'    t1 = Now
'    Do
'        DoEvents
'        If ConnectCK = 1 Then Exit Do            '受信成功
'        If ConnectCK = 2 Then Exit Do            'ビジー
'        If ConnectCK = 3 Then Exit Do            '回線エラー
''        If ConnectCK = 4 Then Exit Do            '受信失敗
''        If ConnectCK = 5 Then Exit Do            '警報通知
'        If ConnectCK = 6 Then Exit Do            '応答がありません。
'
'        t2 = Now
'
'        '５分待ってもイベントが起きなければ電話を切る。
'        If DateDiff("s", DateAdd("s", 600, t1), t2) > 0 Then Exit Do
'    Loop
'
'        StatusBar1.Panels(1).Text = "回線切断開始..."
'        RasClient1.HangUp  '電話を切る
'        StatusBar1.Panels(1).Text = "回線切断..."
'        Call Sleep(2000)
'        ConnectCK = 0
'        StatusBar1.Panels(1).Text = ""
'End Sub

Private Function keihouCK() As Boolean
    Dim f1 As Integer
    Dim f2 As Integer
    Dim sa1 As Variant
    Dim sa2 As Variant
    Dim bf1 As String
    Dim bf2 As String
    Dim i As Integer
    Dim j As Integer
    
    Dim strm As New ADODB.Stream
    Dim bf As String
    Dim sa As Variant
    Dim f As Integer
    
    Dim kfs As String
    Dim kf As Boolean
    kf = False
    Dim kff As Boolean
    kff = False
    Dim st As Integer
'警報テキストファイルが存在するか
    keihouCK = False
    
    '傾斜計
    Dim cp As Integer
    Dim maxTen As Integer
    Dim k As Integer
    Dim cb As String
    
    For i = 1 To keisyaCo
        
        f = FreeFile
        Open g_kankyoPath & g_keisyaConf(i) For Input As #f
        Input #f, cp
        maxTen = cp
        k = 0
        j = 0
        Do While Not (EOF(f))
            j = j + 1
            Line Input #f, cb
            If InStr(";:#", Left$(cb, 1)) = 0 Then
                k = k + 1
                sa = Split(cb, ",")
                g_keisyaDepth(k) = sa(1)
    '            Debug.Print k, keisyaDepth(k)
            End If
        Loop
        Close #f
            
        f1 = FreeFile
        Open KEISOKU.keihou_path & keisyaR(i) For Input As #f1
        Line Input #f1, bf1
        Close #f1
            
        With strm
          'キャラクタコードを設定
          .Charset = "UTF-8"
          'Streamオブジェクトを開く
          .Open
          'テキストファイルをStreamオブジェクトに読み込み
          .LoadFromFile KEISOKU.keihou_path & keisyaKanri(i)
          'Streamオブジェクトの内容を変数strDataに読み込み
          bf = .ReadText
          '変数strDataの内容をイミディエイトウィンドウに出力
          'Debug.Print bf2
          'Streamオブジェクトを閉じる
          .Close: Set strm = Nothing
        End With
        
        sa = Split(bf, vbCrLf)
        bf2 = sa(0)
'        f1 = FreeFile
'        Open KEISOKU.keihou_path & keisyaKanri(i) For Input As #f1
'        Line Input #f1, bf2
'        Close #f1
        
        sa1 = Split(bf1, ",")
        sa2 = Split(bf2, ",")
        
        For j = 1 To UBound(sa2)
            If sa1(j) <> 999999 Then
                If Val(sa2(j)) <= Val(sa1(j)) Then
                    If sa2(j) <> 999999 Then
                        If kf = False Then
        '                    Debug.Print sa1(0)
        '                    Debug.Print keisyaName(i)
                            kfs = sa1(0) & vbCrLf & keisyaName(i) & vbCrLf
                            kf = True
                            kff = True
                        End If
        '                Debug.Print "深度" & g_keisyaDepth(j) & "m : ", sa1(j), "mm"
                        kfs = kfs & "深度" & g_keisyaDepth(j) & "m : " & sa1(j) & "mm" & vbCrLf
                    End If
                End If
            End If
        Next j
        kf = False
    Next i
    
    '水位計
    For i = 1 To SuiiCo
        f1 = FreeFile
        Open KEISOKU.keihou_path & Suii(i) For Input As #f1
        Line Input #f1, bf1
        Close #f1
            
        With strm
          'キャラクタコードを設定
          .Charset = "UTF-8"
          'Streamオブジェクトを開く
          .Open
          'テキストファイルをStreamオブジェクトに読み込み
          .LoadFromFile KEISOKU.keihou_path & SuiiKanri(i)
          'Streamオブジェクトの内容を変数strDataに読み込み
          bf = .ReadText
          '変数strDataの内容をイミディエイトウィンドウに出力
          'Debug.Print bf2
          'Streamオブジェクトを閉じる
          .Close: Set strm = Nothing
        End With
        
        sa = Split(bf, vbCrLf)
        bf2 = sa(0)
'        f1 = FreeFile
'        Open KEISOKU.keihou_path & keisyaKanri(i) For Input As #f1
'        Line Input #f1, bf2
'        Close #f1
        
        sa1 = Split(bf1, ",")
        sa2 = Split(bf2, ",")
        
        'For j = 1 To UBound(sa2)
         If sa1(i) <> 999999 Then
            If Val(sa2(2)) <= Val(sa1(i)) And Val(sa1(i)) <= Val(sa2(1)) Then
            Else
'                Debug.Print sa1(0)
'                Debug.Print "内水位(m)", sa1(i)
                If kf = False Then
                    kfs = kfs & vbCrLf
                    kf = True
                End If
                If Val(sa1(i)) <> 999999 Then
                    kfs = kfs & sa1(0) & vbCrLf & "内水位(m) : " & sa1(i) & vbCrLf
                    kff = True
                End If
            End If
        End If
        'Next j
        kf = False
    Next i
    
    '切梁軸力
    Dim jj As Integer
    For i = 1 To KiriBCo
        f1 = FreeFile
        Open KEISOKU.keihou_path & kiribari(i) For Input As #f1
        Line Input #f1, bf1
        Close #f1
            
        With strm
          'キャラクタコードを設定
          .Charset = "UTF-8"
          'Streamオブジェクトを開く
          .Open
          'テキストファイルをStreamオブジェクトに読み込み
          .LoadFromFile KEISOKU.keihou_path & kiribariKanri(i)
          'Streamオブジェクトの内容を変数strDataに読み込み
          bf = .ReadText
          '変数strDataの内容をイミディエイトウィンドウに出力
          'Debug.Print bf2
          'Streamオブジェクトを閉じる
          .Close: Set strm = Nothing
        End With
        
        sa = Split(bf, vbCrLf)
        bf2 = sa(0)
'        f1 = FreeFile
'        Open KEISOKU.keihou_path & keisyaKanri(i) For Input As #f1
'        Line Input #f1, bf2
'        Close #f1
        
        sa1 = Split(bf1, ",")
        sa2 = Split(bf2, ",")
        
        For j = 1 To UBound(sa2)
            If sa1(j) <> 999999 Then
                If Val(sa1(j)) <= Val(sa2(j)) Then
                Else
    '                Debug.Print sa1(0)
    '                Debug.Print "内水位(m)", sa1(i)
                    If kf = False Then
                        kfs = kfs & vbCrLf
                        kf = True
                    End If
                    If Val(sa1(j)) <> 999999 Then
                        Select Case j
                        Case 1:  kfs = kfs & sa1(0) & " 切梁 1 段目 3通: " & sa1(j) & vbCrLf
                        Case 2:  kfs = kfs & sa1(0) & " 切梁 1 段目 C通: " & sa1(j) & vbCrLf
                        Case 3:  kfs = kfs & sa1(0) & " 切梁 2 段目 3通: " & sa1(j) & vbCrLf
                        Case 4:  kfs = kfs & sa1(0) & " 切梁 2 段目 C通: " & sa1(j) & vbCrLf
                        Case 5:  kfs = kfs & sa1(0) & " 切梁 3 段目 3通: " & sa1(j) & vbCrLf
                        Case 6:  kfs = kfs & sa1(0) & " 切梁 3 段目 C通: " & sa1(j) & vbCrLf
                        Case 7:  kfs = kfs & sa1(0) & " 切梁 4 段目 3通: " & sa1(j) & vbCrLf
                        Case 8:  kfs = kfs & sa1(0) & " 切梁 4 段目 C通: " & sa1(j) & vbCrLf
                        Case 9:  kfs = kfs & sa1(0) & " 切梁 5 段目 3通: " & sa1(j) & vbCrLf
                        Case 10: kfs = kfs & sa1(0) & " 切梁 5 段目 C通: " & sa1(j) & vbCrLf
                        End Select
                        kff = True
                    End If
                End If
            End If
        Next j
    Next i
    
'    Debug.Print kfs
    
    If kff = True Then
        f = FreeFile
        Open App.Path & "\" & KeihouFile For Output As #f
        Print #f, kfs;
        Print #f, "--------"
        Close #f
    End If
    
    Dim kPath As String
    kPath = App.Path & "\" & KeihouFile
    If FileExists(kPath) = True Then
        keihouCK = True
        
'        ここで､"Send000.txt"の内容をチェックして
'        中身が空だったら keihouCK = False で戻る
'        Call msgCK(st, kPath)
'        If st = 0 Then
'            Call DelFile(kPath)
'            keihouCK = False
'        Else
'            keihouCK = True
'        End If
        
    End If
End Function

Private Sub msgCK(st As Integer, kPath As String)
    Dim bf As String
    Dim f As Integer
    f = FreeFile
    Open kPath For Input As #f
    Line Input #f, bf
'    Line Input #f, bf
    Close #f
    If Trim(bf) = "" Then
        st = 0
    Else
        st = -1
    End If
End Sub

Private Sub FileJoinSend()
'////////////////////////////////////////////////////////////////
'ファイルを連結して、送信メッセージテキストを作成 & 送信

    Dim i As Integer, f As Integer
    Dim f2 As Integer
    Dim f3 As Integer
    Dim L As String
    Dim SS1$
    
    Dim ntm As Date
       
'    On Error GoTo FileJoinSend9999
    strData = "" '"以下のデータが管理値をオーバーしています。" & vbCrLf
    f3 = FreeFile
    Open App.Path & "\" & KeihouFile For Input As #f3
    Do While Not (EOF(f3))
        Line Input #f3, L
        'If Left$(L, 1) <> ";" Then
            strData = strData & L & vbCrLf
        'End If
    Loop
    Close #f3
    
'    If MailTabl.JyusinSW = 1 Then
'        Call MailRead(MailTabl) 'メール受信を実行
'    End If
    
'Exit Sub
    
    Call MailSend(MailTabl) 'メール送信を実行
    ntm = Now
    Call WriteIni("メール送信", "最終送信日時", Format$(ntm, "yyyy/mm/dd hh:nn:ss"), mINIfile)
'    Label2(2) = Format$(ntm, "yyyy/mm/dd hh:nn:ss")
    
    strData = ""
    
    StatusBar1.Panels(1).Text = ""
Exit Sub
FileJoinSend9999:
    Call wait3(1000)
    Resume
End Sub

Private Sub MailSend(MailTbl As MailType)
    Dim i As Integer
    With MailTbl
        
        .SendCO = CInt(GetIni("メール送信", "送信数", mINIfile))
        For i = 1 To .SendCO
            .SendName(i) = GetIni("メール送信", "送信先" & CStr(i), mINIfile)
        Next i
    End With
    
    Dim ssb As String
    ssb = (GetIni("メール送信", "subject", mINIfile))
    
    Dim MailSmtpServer As String        ' SMTPサーバ
    Dim MailFrom As String              ' 発信者
    Dim MailTo As String                ' 宛先
    Dim MailToBCC As String             ' 宛先(BCC)
    Dim MailSubject As String           ' 件名
    Dim MailBody As String              ' 本文
    Dim MailAddFile As String           ' 添付ファイル名(1ファイルのみ対応)
    Dim strMSG As String                ' 結果メッセージ
    Dim strMSG2 As String               ' 結果メッセージのWORK
    
'    MailSmtpServer = GetIni("メール送信", "サーバー名", mINIfile)
'    MailFrom = GetIni("メール送信", "メールアドレス", mINIfile)
    
    MailSmtpServer = "smtp.gmail.com"       'MailTabl.ServerName      ' SMTPサーバ
    MailFrom = "<atic.alertmail@gmail.com>"   'MailTabl.ClientMailAddress            ' 発信者
    MailTo = MailTabl.SendName(1)           ' 宛先
    MailSubject = ssb                       ' 件名
    MailBody = strData                      ' 本文
    MailAddFile = ""                        ' 添付ファイル名(1ファイルのみ対応)
    'strMSG                ' 結果メッセージ
    'strMSG2               ' 結果メッセージのWORK
    
    Dim szTo As String
    If MailTabl.SendCO > 1 Then
        szTo = ""
        For i = 2 To MailTabl.SendCO
            szTo = szTo & "," & MailTabl.SendName(i)
        Next i
        MailToBCC = szTo
    End If

    'strMSG2 = SendMailByCDO(MailSmtpServer, MailFrom, MailTo, "", "", _
        MailSubject, MailBody, MailAddFile)
    ' 文字コードを指定する場合は以下のように変更します。(ISO2022JPの例)
        strMSG2 = SendMailCDO(RAS(1).eName, MailSmtpServer, MailFrom, MailTo, "", MailToBCC, _
        MailSubject, MailBody, MailAddFile, cdoISO_2022_JP)
    '-----------------------------------------------------------------------
    ' 送信不成功の場合はエラーメッセージを集積
    If strMSG2 <> "OK" Then
        If strMSG <> "" Then strMSG = strMSG & vbCr
        strMSG = strMSG & strMSG2 & " (" & MailTo & ")"
        Call WriteLog("メール送信 失敗")
    Else
        Call WriteLog("メール送信 終了")
    End If

End Sub

Private Sub SendLogMove()
    Dim f As Integer
    Dim f2 As Integer
    Dim sc As Integer
    Dim bf As String
    
    sc = GetIni("メール送信", "sc", mINIfile)
    sc = sc + 1
    f = FreeFile
    Open App.Path & "\" & KeihouFile For Input As #f
    f2 = FreeFile
    Open App.Path & "\sendlog\send" & Format(sc, "0000") & ".txt" For Output As #f2
    Do While Not (EOF(f))
        Line Input #f, bf
        Print #f2, bf
    Loop
    Close #f2
    Close #f
    Call WriteIni("メール送信", "sc", (sc), mINIfile)

    Call DelFile(App.Path & "\" & KeihouFile)
End Sub


'Private Sub RasInitial(ByVal id As Integer)
'    'RASエントリーから目的の接続先を探す
'    Dim i As Integer
'
'    With MainForm
'    If .RasClient1.ItemCount = 0 Then
'        'エントリー無し
'        MsgBox "エントリー無し": End 'Stop
'    Else
'            For i = 0 To .RasClient1.ItemCount - 1
'                If .RasClient1.Item(i) = RAS(id).eName Then
'                    'エントリー番号を保持
'                    RAS(id).iEntNo = i
'                    Exit For
'    '            Else
'    '                '接続先が見つからない
'    '                MsgBox "接続先が見つからない": End 'Stop
'                End If
'            Next i
'            If i = RasClient1.ItemCount Then
'                MsgBox "接続先 " & RAS(id).eName & " が見つからない": End
'            End If
'    End If
'    End With
'End Sub

Private Sub waits(se As Double)
    Dim Waittime As Date
    Dim Ntime As Date
    Waittime = DateAdd("s", se, Now)
    Ntime = Now
    Do Until Now >= Waittime
        DoEvents
        If Ntime <> Now Then
            Label3 = Format$(Ntime, "yyyy/mm/dd hh:mm:ss")
            Ntime = Now
        End If
    Loop
End Sub

Private Sub wait3(tm As Long)
    Dim i As Long
    i = 0
    Do
        DoEvents
        Call Sleep(10)
        i = i + 10
    Loop While i < tm
End Sub


