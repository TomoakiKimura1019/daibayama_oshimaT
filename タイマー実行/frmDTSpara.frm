VERSION 5.00
Object = "{2641A793-B080-4A11-96D3-BA6820C8A647}#4.2#0"; "xDateN.ocx"
Object = "{9563DB83-ADCE-4722-A569-DEF0B7A18131}#4.2#0"; "xTimeN.ocx"
Begin VB.Form frmDTSpara 
   Caption         =   "データシートパラメータ設定"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   Icon            =   "frmDTSpara.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "表示"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin xDateNLib.xDateN xDateN1 
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1215
         _Version        =   262146
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   -2147483640
         AlignmentH      =   1
         BorderStyle     =   4
         DateFormat      =   3
         Value           =   41479
         YearValue       =   2013
         MonthValue      =   7
         DayValue        =   24
         Week            =   4
         DisplayString   =   "2013/07/24"
         SpinButton      =   0   'False
         FocusBorder     =   1
         ZeroFill        =   -1  'True
         Caption         =   "xDateN1"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   0
         InputAreaBorder =   0
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
         PMenuCaption6   =   "書式コピー(&F)"
         PEnabled6       =   -1  'True
         PHidden6        =   0   'False
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
      Begin xDateNLib.xDateN xDateN1 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   1215
         _Version        =   262146
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         DateFormat      =   3
         Value           =   41479
         YearValue       =   2013
         MonthValue      =   7
         DayValue        =   24
         Week            =   4
         DisplayString   =   "2013/07/24"
         SpinButton      =   0   'False
         FocusBorder     =   1
         ZeroFill        =   -1  'True
         Caption         =   "xDateN1"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   0
         InputAreaBorder =   0
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
         PMenuCaption6   =   "書式コピー(&F)"
         PEnabled6       =   -1  'True
         PHidden6        =   0   'False
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
      Begin xTimeNLib.xTimeN xTimeN1 
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   705
         _Version        =   262146
         _ExtentX        =   1244
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         DisplayItem     =   1
         TimeFormat      =   4
         Value           =   0.602083333333333
         HourValue       =   14
         MinuteValue     =   27
         SecondValue     =   0
         DisplayString   =   "14:27"
         SpinButton      =   0   'False
         FocusBorder     =   1
         SecondInput     =   0   'False
         Caption         =   "xTimeN1"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   0
         InputAreaBorder =   0
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
         PMenuCaption6   =   "書式コピー(&F)"
         PEnabled6       =   -1  'True
         PHidden6        =   0   'False
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
      Begin xTimeNLib.xTimeN xTimeN1 
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   705
         _Version        =   262146
         _ExtentX        =   1244
         _ExtentY        =   661
         _StockProps     =   79
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         DisplayItem     =   1
         TimeFormat      =   4
         Value           =   0.602083333333333
         HourValue       =   14
         MinuteValue     =   27
         SecondValue     =   0
         DisplayString   =   "14:27"
         SpinButton      =   0   'False
         FocusBorder     =   1
         SecondInput     =   0   'False
         Caption         =   "xTimeN1"
         CaptionRatio    =   0
         InputAreaBorder =   0
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
         PMenuCaption6   =   "書式コピー(&F)"
         PEnabled6       =   -1  'True
         PHidden6        =   0   'False
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
      Begin VB.Label Label4 
         Caption         =   "終了日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "開始日時"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   450
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDTSpara"
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
        ckdate1 = Format(xDateN1(0).Value, "yyyy/mm/dd") & " " & xTimeN1(0).Value
        ckdate2 = Format(xDateN1(1).Value, "yyyy/mm/dd") & " " & xTimeN1(1).Value
        If ckdate1 > ckdate2 Then
            MsgBox "日付指定の誤り", vbCritical, "エラーメッセージ"
            xDateN1(0).SetFocus
            Exit Sub
        End If
        Hsheet.SD = ckdate1
        Hsheet.ED = ckdate2
        Unload Me
        Call frmDts.Sakuhyou
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Top = frmSakuzu.Top + 1000
    Left = frmSakuzu.Left + 500
    SD = Hsheet.SD
    ED = Hsheet.ED

    xDateN1(0).Value = DateValue(SD)
    xTimeN1(0).Value = TimeValue(SD)
    xDateN1(1).Value = DateValue(ED)
    xTimeN1(1).Value = TimeValue(ED)

End Sub


