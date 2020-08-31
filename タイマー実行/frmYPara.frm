VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Begin VB.Form frmYPara 
   Caption         =   "Ｙ軸パラメータ設定"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   Icon            =   "frmYPara.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Frame Frame3 
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
         Left            =   960
         TabIndex        =   6
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
         Left            =   960
         TabIndex        =   7
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
         Left            =   960
         TabIndex        =   8
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
         Left            =   360
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
         Left            =   360
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
         Left            =   360
         TabIndex        =   3
         Top             =   390
         Width           =   600
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表示"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frmYPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SD As Date, ED As Date, xBUN As Integer, mkMAX As Integer
Dim YMAX(3) As Single, YMIN(3) As Single, yBUN(3) As Integer

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
        
        Kkeiji.YMIN = CSng(xTextN1(0).Text)
        Kkeiji.YMAX = CSng(xTextN1(1).Text)
        Kkeiji.yBUN = CInt(xTextN1(2).Text)
        
        Call WriteIni("常時表示設定", "Ｙ軸最小値", CStr(Kkeiji.YMIN), CurrentDir & "計測設定.ini")
        Call WriteIni("常時表示設定", "Ｙ軸最大値", CStr(Kkeiji.YMAX), CurrentDir & "計測設定.ini")
        Call WriteIni("常時表示設定", "Ｙ軸分割数", CStr(Kkeiji.yBUN), CurrentDir & "計測設定.ini")
        Unload Me
        Call KeijiInit
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, L As String
    Dim ss As String

    Top = MainForm.Top + 700
    Left = MainForm.Left + 1500
    
    YMIN(1) = Kkeiji.YMIN
    YMAX(1) = Kkeiji.YMAX
    yBUN(1) = Kkeiji.yBUN
    
    xTextN1(0).Text = YMIN(1)
    xTextN1(1).Text = YMAX(1)
    xTextN1(2).Text = yBUN(1)
    
End Sub

Private Sub xTextN1_LostFocus(Index As Integer)
    If IsNumeric(xTextN1(Index).Text) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        Exit Sub
    End If
    If Index = 2 And CInt(xTextN1(Index).Text) <= 0 Then
        MsgBox "分割数の入力に誤りがあります。もう一度設定してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        Exit Sub
    End If
End Sub


