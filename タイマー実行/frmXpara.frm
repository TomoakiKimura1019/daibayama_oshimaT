VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Begin VB.Form frmXpara 
   Caption         =   "日付設定"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "frmXpara.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      Begin xTextNLib.xTextN xTextN1 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   615
         _Version        =   262145
         _ExtentX        =   1085
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
      Begin VB.Label Label2 
         Caption         =   "（最大数＝10）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "分割数"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   5
         Top             =   915
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  '右揃え
         Caption         =   "表示日数"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "（最大日数＝200）"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   3
         Top             =   435
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表示"
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmXpara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SD As Date, ED As Date
Dim ds1 As Object

Private Sub Command1_Click(Index As Integer)
Dim ckdate1 As Date, ckdate2 As Date
    If Index = 0 Then
        Kkeiji.Xmax = xTextN1(0).Text
        Kkeiji.xBUN = xTextN1(1).Text
        Call WriteIni("常時表示設定", "Ｘ軸日数", CStr(Kkeiji.Xmax), CurrentDir & "計測設定.ini")
        Call WriteIni("常時表示設定", "Ｘ軸分割数", CStr(Kkeiji.xBUN), CurrentDir & "計測設定.ini")
        Unload Me
        Call KeijiInit
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Top = MainForm.Top + 1000
    Left = MainForm.Left + 500
    
    xTextN1(0).Text = Format$(Kkeiji.Xmax, "0")
    xTextN1(1).Text = Format$(Kkeiji.xBUN, "0")
End Sub

Private Sub xTextN1_LostFocus(Index As Integer)
    If IsNumeric(xTextN1(Index).Text) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        GoTo txt_skip
    End If
    If xTextN1(Index).Text <= 0 Then
        MsgBox "0より大きい数値を入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        GoTo txt_skip
    End If
    If Index = 0 And xTextN1(Index).Text > 180 Then
        MsgBox "最大日数より小さい数値を入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        GoTo txt_skip
    End If
    If Index = 1 And xTextN1(Index).Text > 10 Then
        MsgBox "最大値より小さい数値を入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
        GoTo txt_skip
    End If
txt_skip:
End Sub


