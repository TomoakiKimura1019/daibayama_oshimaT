VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "2"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4350
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7620
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4920
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
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   1
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   870
         Width           =   735
      End
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
      TabIndex        =   7
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
      TabIndex        =   6
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
'    If App.PrevInstance Then
'      Unload Me
'    End If
    
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
    
'    Me.Height = 1545 '2865
    Me.Width = 4545
    
    Caption = GetIni("現場名", "現場名", mINIfile) + " データ送信"
    
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    Timer1.Enabled = True
    
'    NoRAS = GetIni("RAS", "NoRAS", mINIfile)
    
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
        Open CurrentDir & App.EXEName & "-event.log" For Append Lock Write As #f
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

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Dim Tstime As String
    Dim keisoku_f As Boolean
    Dim ret As Integer
    
    Tstime = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    
    If SoushinFileCK = False Then Call tmSub
    
    Timer1.Enabled = True
    
    Unload Me
    
End Sub

Private Sub tmSub()
    
    Dim ii As Integer
    
    '警報データファイルがあったら、そのデータを送信
'    keihouF = keihouCK
    
    '計測データフォルダの中身があったら、そのデータを送信
    For ii = 1 To fco
        If CheckDataFile(TdsDataPath(ii)) <> 0 Then
            StatusBar1.Panels(1).Text = "送信開始-" & ii
            waits 5  ' 5秒待つ
            SendDataF = True
            
            If SendDataF = True Then
                Call FindDataFile(2, TdsDataPath(ii), ii)
                SendDataF = False
            End If
            StatusBar1.Panels(1).Text = ""
        
        End If
    Next ii

Exit Sub

End Sub

Private Function keihouCK() As Boolean
'警報テキストファイルが存在するか
    keihouCK = False
    Dim kPath As String
    kPath = KEISOKU.keihou_path & "Send000.txt"
    If FileExists(kPath) = True Then
        keihouCK = True
'        iDaialID = 2
'        Call Soushin(iDaialID)
        'Call FileJoinSend
        'Call SendLogMove
    End If
End Function

Private Sub SendLogMove()
    Dim f As Integer
    Dim f2 As Integer
    Dim sc As Integer
    Dim bf As String
    
    sc = GetIni("メール送信", "sc", mINIfile)
    sc = sc + 1
    f = FreeFile
    Open KEISOKU.keihou_path & "send000.txt" For Input As #f
    f2 = FreeFile
    Open App.Path & "\sendlog\send" & Format(sc, "0000") & ".txt" For Output As #f2
    Do While Not (EOF(f))
        Line Input #f, bf
        Print #f2, bf
    Loop
    Close #f2
    Close #f
    Call WriteIni("メール送信", "sc", (sc), mINIfile)

    Call DelFile(KEISOKU.keihou_path & "send000.txt")
End Sub

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

