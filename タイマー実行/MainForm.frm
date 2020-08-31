VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Form3"
   ClientHeight    =   2235
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows の既定値
   Visible         =   0   'False
   Begin VB.Timer wTimer 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Left            =   4440
      Top             =   1200
   End
   Begin VB.CommandButton Command5 
      Caption         =   "手動実行"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   960
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   1935
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8811
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
   Begin xTextNLib.xTextN xTextN2 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
      _Version        =   262145
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      AlignmentH      =   1
      Caption         =   "xTextN1"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      Locked          =   -1  'True
      KillClickFocus  =   -1  'True
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
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
      _Version        =   262145
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      Text            =   "2013/10/10 12:00:00"
      AlignmentH      =   1
      Caption         =   "xTextN1"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      Locked          =   -1  'True
      KillClickFocus  =   -1  'True
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
      Left            =   1320
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1935
      _Version        =   262145
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      AlignmentH      =   1
      Caption         =   "xTextN1"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      Locked          =   -1  'True
      KillClickFocus  =   -1  'True
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
   Begin xTextNLib.xTextN xTextNtime 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
      _Version        =   262145
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   77
      Text            =   "2013/10/10 12:00:00"
      AlignmentH      =   1
      Caption         =   "xTextN1"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      Locked          =   -1  'True
      KillClickFocus  =   -1  'True
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
   Begin VB.Label Label8 
      Alignment       =   1  '右揃え
      Caption         =   "現在時刻"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "インターバル"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "前回測定時間"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "次回測定時間"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' SetWindowPos 関数
Private Declare Function SetWindowPos Lib "USER32.DLL" ( _
    ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long

' 定数の定義
Private Const HWND_TOPMOST   As Long = -1     ' 最全面に表示する
Private Const HWND_NOTOPMOST As Long = -2     ' 最前面に表示するのをやめる
Private Const SWP_NOSIZE     As Long = &H1    ' サイズを変更しない
Private Const SWP_NOMOVE     As Long = &H2    ' 位置を変更しない

Dim STtime As Date, minTIME As Date
Dim keisoku_f As Boolean
Dim Thistime As String
Dim Stat As Integer

Dim FilePath() As String
Dim FileCo As Integer
Dim pFileName() As String

Private nidSysInfo As NOTIFYICONDATA
Private lRetVal As Long

Dim TimEvent As Boolean
Dim fZorder As Integer

Private Sub Form_Resize()
    Dim zo As Integer
    Dim tmp As String
    tmp = GetIni("Form", "Zorder", App.Path & "\ExecTimer.ini")
    
    If tmp = "-1" Then
        ' このフォームを常に最前面に表示する (サイズと位置は変更しない)
        Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        fZorder = -1
    Else
        ' 解除したい場合
        Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
        fZorder = 0
    End If
End Sub

Private Sub Label3_DblClick()
    Call SetWindowPos(MainForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Timer1.Enabled = False
    ivCode = 1
    Form4.Show vbModal
    Call IntvWrite
    Call DayTimeWrite
    If fZorder = -1 Then
        Call SetWindowPos(MainForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Timer1.Enabled = True
End Sub

Private Sub Label6_DblClick()
    Call SetWindowPos(MainForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Timer1.Enabled = False
    ivCode = 0
    Form4.Show vbModal
    Call IntvWrite
    Call DayTimeWrite
    If fZorder = -1 Then
        Call SetWindowPos(MainForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Timer1.Enabled = True
End Sub

Private Sub Label7_DblClick()
    Call SetWindowPos(MainForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Timer1.Enabled = False
    ivCode = 1
    Form3.Show vbModal
    Call IntvWrite
    Call DayTimeWrite
    If fZorder = -1 Then
        Call SetWindowPos(MainForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Timer1.Enabled = True
End Sub

Private Sub Label9_DblClick()
    Call SetWindowPos(MainForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Timer1.Enabled = False
    ivCode = 0
    Form3.Show vbModal
    Call IntvWrite
    Call DayTimeWrite
    If fZorder = -1 Then
        Call SetWindowPos(MainForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    End If
    Timer1.Enabled = True
End Sub

Private Sub Timer2_Timer()
    xTextNtime.Text = Format$(Now, "YYYY/MM/DD hh:mm:ss")
End Sub

Private Sub Timer1_Timer()
    Call tmSub
End Sub

Private Sub Command5_Click()
'手動計測ボタン

    Dim j As Integer
    Dim i As Integer, f As Integer
    Dim ckTIME As Date
    Dim Stat As Integer
    Dim Thistime As String
    Dim Syudou As Boolean
    Dim Z_jidou_time As Date
    Dim minTIME As Date
    Dim t1 As Date, t2 As Date
    
'    On Error GoTo KeiERR

    If vbOK = MsgBox("「OK」をクリックすると、計測します｡", vbOKCancel + vbInformation, "手動計測") Then
        Syudou = True
    Else
        Syudou = False
    End If
    If Syudou = True Then
        StatusBar1.Panels(1).Text = "*** 手動計測中 ***"
        Enabled = False
        
            For j = 1 To FileCo
'                sw = False
'                For i = 0 To List1.ListCount - 1
'                    If UCase(pFileName(j)) = UCase(Mid(List1.List(i), 9)) Then
'                        sw = True
'                        Exit For
'                    End If
'                Next i
'                If sw = False Then
                    Call CPexec(FilePath(j))
'                End If
            Next j
        
        Enabled = True
        
        StatusBar1.Panels(1).Text = ""
    End If
    
    Exit Sub

KeiERR:
    Close
'    Unload Form2
'''    Unload Form1
    
    If Err.Number = 10000 Then
        MsgBox "通信エラー"
    Else
        MsgBox "エラー:" & Err.Number
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer, vTMP
    
    Dim fTop As Long, fLeft As Long
    Dim TMP0 As Variant
    TMP0 = GetIni("Form", "top", CurrentDir & "ExecTimer.ini")
    If TMP0 <> "" Then
        fTop = TMP0
        If fTop < 0 Then fTop = 0
    End If
    TMP0 = GetIni("Form", "left", CurrentDir & "ExecTimer.ini")
    If TMP0 <> "" Then
        fLeft = TMP0
        If fLeft < 0 Then fLeft = 0
    End If
    
    Top = fTop
    Left = fLeft
    Dim cp As String
    cp = GetIni("SYSTEM", "caption", CurrentDir & "ExecTimer.ini")
    Caption = cp & " タイマー起動"
    
    Dim rc As Integer
    Dim t1 As Date, t2 As Date
'    Dim i As Integer

    '2001/05/16
'    t1 = Timer
'    frmInitMsg.Show: frmInitMsg.Refresh
'    If Command$ = "TEST" Then
'        Do
'            DoEvents
'            t2 = Timer
'            If (t2 - t1) > 2 Then Exit Do
'        Loop
'    Else
'    End If
'    Unload frmInitMsg

    Me.Enabled = False
    
    Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
    '次回計測時間計算
'    Debug.Print toTMSstring(KE_intv)
    Call IntvWrite
    xTextN2(0).Text = Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN2(1).Text = toTMSstring(KE_intv)
    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
        
    With nidSysInfo
        .cbSize = Len(nidSysInfo)
        .hWnd = Me.hWnd
        .uID = 1
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MBUTTONDOWN    'Calllback Message
        .hIcon = Me.Icon
        .szTip = "起動監視" & vbNullChar
    End With
    
    FileCo = Val(GetIni("system", "FileCo", CurrentDir & "ExecTimer.ini"))
    ReDim pFileName(FileCo)
    
    For i = 1 To FileCo
        ReDim Preserve FilePath(i)
        FilePath(i) = (GetIni("system", "File" & i, CurrentDir & "ExecTimer.ini"))
        pFileName(i) = GetFullPasToFileName(FilePath(i))
    Next i
    
    Me.Enabled = True
    
    '自動計測開始
    keisoku_f = False
'    Me.SetFocus
    Timer1.Interval = 200
    Timer1.Enabled = True
    Timer2.Interval = 250
    Timer2.Enabled = True
    
    wTimer.Enabled = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Ret As Long
    Dim RetString As String
    Dim i As Integer, ENDsw As Boolean, f As Integer
    Dim rc As Integer
    
    On Error Resume Next
    
    If UnloadMode < 1 Then
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
        
        Call WriteIni("Form", "top", (Top), CurrentDir & "ExecTimer.ini")
        Call WriteIni("Form", "left", (Left), CurrentDir & "ExecTimer.ini")
       
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

Private Sub tmSub()
    Dim i As Integer, j As Integer
    Dim t1 As Date
 
        MDY = Now
        Thistime = Format$(MDY, "yyyy/mm/dd hh:nn:ss")
        xTextNtime.Text = Format$(Thistime, "YYYY/MM/DD hh:mm:ss")
        
        If Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss") = Thistime Then keisoku_f = True
        
        If keisoku_f = True Then
            MainForm.StatusBar1.Panels(1).Text = "*** 実行中 ***"
            MainForm.Enabled = False
                
'            Form2.Show
            For j = 1 To FileCo
'                sw = False
'                For i = 0 To List1.ListCount - 1
'                    If UCase(pFileName(j)) = UCase(Mid(List1.List(i), 9)) Then
'                        sw = True
'                        Exit For
'                    End If
'                Next i
'                If sw = False Then
                    Call CPexec(FilePath(j))
                    WaitTime 2000
'                End If
            Next j
            
'            Unload Form2
            
            
            '次回計測時間計算
            If keisoku_f = True Then
                keisoku_f = False
                Keisoku_TimeZ = Keisoku_Time
                Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
            End If
            Call IntvWrite

            Call DayTimeWrite
            
            MainForm.Enabled = True
            If Stat = 0 Then MainForm.StatusBar1.Panels(1).Text = ""
        End If
            
99      '計測時間がすぎた場合
        If DateDiff("s", Keisoku_Time, MDY) >= 0 Then  'If nt < Now Then
            Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
            Call IntvWrite
            
            xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            
            'ログ
            Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 計測時間が過ぎていたため、再設定しました。")
        End If

End Sub

Private Sub DayTimeWrite()
    xTextN2(0).Text = Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN2(1).Text = toTMSstring(KE_intv)
    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")

End Sub

Private Sub CPexec(fi As String)
    Dim pa0 As String
    Dim pa As String
    Dim nm As String
    pa0 = App.Path
    
    pa = GetPathNameToFullPas(fi)
    nm = GetFullPasToFileName(fi)
    
    SetCurrentDirectory pa & "\"

    ' 子プロセス起動
    Dim lngPid As Long
'    lngPid = CLng(Shell(fi, vbHide))
    lngPid = CLng(Shell(fi, vbNormalFocus))
    
    SetCurrentDirectory pa0 & "\"
    
End Sub

Private Sub wTimer_Timer()
   TimEvent = True
End Sub

'
' 特定の時間待ちをする
'
'   Ti = 待ち時間　(ms)
'
Public Sub WaitTime(Ti As Single)

   wTimer.Enabled = False
   wTimer.Interval = Ti
   TimEvent = False
   wTimer.Enabled = True
   
   Do While TimEvent = False
       DoEvents
   Loop
   
End Sub


