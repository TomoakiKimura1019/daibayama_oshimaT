VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Begin VB.Form frmBunpuPara 
   Caption         =   "計測値図パラメータ設定"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frmBunpuPara.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Frame Frame1 
      Caption         =   "計測データ日時一覧"
      ForeColor       =   &H00800000&
      Height          =   5175
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   3135
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   240
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmBunpuPara.frx":0442
         Left            =   240
         List            =   "frmBunpuPara.frx":0444
         TabIndex        =   7
         Top             =   675
         Width           =   2655
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ｙ軸設定"
      ForeColor       =   &H00800000&
      Height          =   1575
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin xTextNLib.xTextN xTextN1 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InputAlignmentH =   1
         AlignmentH      =   1
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         TreatEnterAsTab =   -1  'True
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
      Begin xTextNLib.xTextN xTextN1 
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   660
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InputAlignmentH =   1
         AlignmentH      =   1
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         TreatEnterAsTab =   -1  'True
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
      Begin xTextNLib.xTextN xTextN1 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   735
         _Version        =   262145
         _ExtentX        =   1296
         _ExtentY        =   661
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         InputAlignmentH =   1
         AlignmentH      =   1
         Caption         =   "xTextN1"
         CaptionRatio    =   0
         TreatEnterAsTab =   -1  'True
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
      Begin VB.Label Label1 
         Caption         =   "分割数"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "最大値"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   765
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "最小値"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表示"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
End
Attribute VB_Name = "frmBunpuPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SD As Date, ED As Date, xBUN As Integer, mkMAX As Integer
Dim YMAX(3) As Single, YMIN(3) As Single, yBUN(3) As Integer
'作図日時、             指定された個数
Dim HIZUKE As Date
'ロードが完了するとTrue
Dim ckSW As Boolean
Dim FileName  As String

'**********************************************************************************************
'   選択した月の計測日時一覧を再描画がする
'**********************************************************************************************
Private Sub Combo1_Click()
    Dim i As Integer, j As Integer
    Dim L As String
    Dim Dco As Long, Dsw As Boolean
    Dim maxREC As Long, Rco As Long, Rpst1 As Long, Rpst2 As Long, Rsw As Boolean
    Dim days As Date, OldTuki As String
    Dim po As Long
    Dim SD As Date, ED As Date
    
    SD = CDate(Combo1.List(Combo1.ListIndex))
    ED = DateAdd("m", 1, SD)
    
    po = STARTpoint(SD)
    
    '計測日時のカウント
    Dco = 0
    
    Screen.MousePointer = 11
    ckSW = False
    List1.Visible = False
    List1.Clear
    
    Open FileName For Input Shared As #1
    Seek #1, po
        Do While Not (EOF(1))
            Line Input #1, L
            days = CDate(Mid$(L, 1, 19))
            If days < SD Then GoTo skip_1
            If days >= ED Then Exit Do
            
            List1.AddItem Mid$(L, 1, 19)

            Dco = Dco + 1
            If DateDiff("s", List1.List(Dco - 1), HIZUKE) = 0 Then List1.ListIndex = Dco - 1
skip_1:
        Loop
    Close #1
    
    ckSW = True
    List1.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim ckdate1 As Date, ckdate2 As Date
    Dim ss As String
    Dim i As Integer

    If Index = 0 Then
        If CSng(xTextN1(0).Text) > CSng(xTextN1(1).Text) Then
            MsgBox "最小値の方が、最大値より大きい値が入力されています。もう一度設定してください。", vbCritical, "エラーメッセージ"
            xTextN1(0).SetFocus
            Exit Sub
        End If
        If CSng(xTextN1(0).Text) = CSng(xTextN1(1).Text) Then
            MsgBox "最大値と最小値に同じ値が入力されています。もう一度設定してください。", vbCritical, "エラーメッセージ"
            xTextN1(0).SetFocus
            Exit Sub
        End If
        If CInt(xTextN1(2).Text) = 0 Then
            MsgBox "分割数の入力に誤りがあります。もう一度設定してください。", vbCritical, "エラーメッセージ"
            xTextN1(2).SetFocus
            Exit Sub
        End If
        
        Hbunpu.YMIN = CSng(xTextN1(0).Text)
        Hbunpu.YMAX = CSng(xTextN1(1).Text)
        Hbunpu.yBUN = CInt(xTextN1(2).Text)
        Hbunpu.SD = HIZUKE

        Unload Me
        Call frmSakuzu.HeniBunpu(False)
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, L As String
    Dim ss As String

    Dim Dco As Long, Dsw As Boolean
    Dim maxREC As Long, Rco As Long, Rpst1 As Long, Rpst2 As Long, Rsw As Boolean
    Dim days As Date, OldTuki As String
    Dim SD As Date, ED As Date, maxTUKI As Integer
    Dim SeekDay As Date, po As Long
    
    Top = frmSakuzu.Top + 700
    Left = frmSakuzu.Left + 500
    
    YMIN(1) = Hbunpu.YMIN
    YMAX(1) = Hbunpu.YMAX
    yBUN(1) = Hbunpu.yBUN
    
    xTextN1(0).Text = YMIN(1)
    xTextN1(1).Text = YMAX(1)
    xTextN1(2).Text = yBUN(1)
    

    Screen.MousePointer = 11

    FileName = KEISOKU.Data_path & DATA_DAT
    
    HIZUKE = 0
    po = STARTpoint(Hbunpu.SD)
    Open FileName For Input Shared As #1
        Seek #1, po
        Do While Not (EOF(1))
            Line Input #1, L
            days = CDate(Mid$(L, 1, 19))
            If Format(Hbunpu.SD, "yyyy/mm/dd hh:nn:ss") = Format(days, "yyyy/mm/dd hh:nn:ss") Then
                HIZUKE = Hbunpu.SD
            End If
            Exit Do
        Loop
    Close #1
    
    Open FileName For Input Shared As #1
        Line Input #1, L: SD = DateSerial(Mid$(L, 1, 4), Mid$(L, 6, 2), 1)
        If LOF(1) - REC_LEN * 2 > 0 Then
            Seek #1, LOF(1) - REC_LEN * 2
        End If
        Do While Not (EOF(1))
            Line Input #1, L
        Loop
    Close #1
    ED = CDate(Mid$(L, 1, 19))


    '開始から現在までの計測月数
    maxTUKI = DateDiff("m", SD, ED)
    
    For i = 0 To maxTUKI
        SeekDay = DateAdd("m", i, SD)
        po = STARTpoint(SeekDay)
        Open FileName For Input Shared As #1
            Seek #1, po
            Do While Not (EOF(1))
                Line Input #1, L
                days = CDate(Mid$(L, 1, 19))
                If Format(SeekDay, "yyyy/mm") = Format(days, "yyyy/mm") Then
                    Combo1.AddItem Format(days, "yyyy年 m月")
                End If
                Exit Do
            Loop
        Close #1
    Next i
    Screen.MousePointer = 0
    Combo1.ListIndex = 0
End Sub

Private Sub List1_Click()
    HIZUKE = List1.List(List1.ListIndex)
End Sub

Private Sub xTextN2_LostFocus(Index As Integer)
    If Not (IsNumeric(xTextN1(Index).Text)) Then
        MsgBox "数値以外が入力されています。", vbCritical, "入力エラー"
        xTextN1(Index).SetFocus
        Exit Sub
    End If
    If Index = 2 And CInt(xTextN1(Index).Text) <= 0 Then
        MsgBox "分割数の入力に誤りがあります。もう一度設定してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        Exit Sub
    End If
End Sub


