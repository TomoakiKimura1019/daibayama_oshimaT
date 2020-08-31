VERSION 5.00
Object = "{2641A793-B080-4A11-96D3-BA6820C8A647}#4.2#0"; "xDateN.ocx"
Object = "{9563DB83-ADCE-4722-A569-DEF0B7A18131}#4.2#0"; "xTimeN.ocx"
Begin VB.Form Form4 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "Form4"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3930
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton Command1 
      Caption         =   "決定"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "キャンセル"
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin xTimeNLib.xTimeN xTimeN1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      _Version        =   262146
      _ExtentX        =   1931
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
      AlignmentH      =   2
      BorderStyle     =   1
      TimeFormat      =   4
      Value           =   0.583333333333333
      HourValue       =   14
      MinuteValue     =   0
      SecondValue     =   0
      UpDownMode      =   1
      DisplayString   =   "14:00:00"
      CaptionRatio    =   0
      InputAreaBorder =   0
      TreatEnterAsTab =   -1  'True
      LostFocusAtFullInput=   -1  'True
      CancelKey       =   1
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
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
      AlignmentH      =   2
      BorderStyle     =   1
      DateFormat      =   3
      Value           =   41488
      YearValue       =   2013
      MonthValue      =   8
      DayValue        =   2
      Week            =   6
      DisplayString   =   "2013/08/02"
      MaxDate         =   "2039/12/31"
      ZeroFill        =   -1  'True
      CaptionRatio    =   0
      InputAreaBorder =   0
      LostFocusAtFullInput=   -1  'True
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
   Begin VB.Label Label1 
      Caption         =   "日付　　　　　　時刻"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If ivCode = 0 Then
        Keisoku_Time = CDate(xDateN1.Value & " " & xTimeN1.Value)
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Caption = "次回開始日時設定"
            
    If ivCode = 0 Then
        xDateN1.Value = Format(Keisoku_Time, "yyyy/mm/dd")
        xTimeN1.Value = Format(Keisoku_Time, "hh:nn:ss")
    End If
End Sub
