VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Begin VB.Form frmSyokiset 
   Caption         =   "観測初期設定"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmSyokiset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7965
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command4 
         Caption         =   "位置設定"
         Height          =   495
         Left            =   4800
         TabIndex        =   0
         Top             =   2640
         Width           =   1215
      End
      Begin xTextNLib.xTextN xTextN1 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
         _Version        =   262145
         _ExtentX        =   2355
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
         Caption         =   "測点数"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   50
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
         Index           =   8
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "方向角"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｘ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｙ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｚ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "ＨＭ"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Index           =   5
         Left            =   3360
         TabIndex        =   14
         Top             =   1080
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｘ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Index           =   6
         Left            =   3360
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｙ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Index           =   7
         Left            =   3360
         TabIndex        =   16
         Top             =   1800
         Width           =   2655
         _Version        =   262145
         _ExtentX        =   4683
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
         Caption         =   "Ｚ座標"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionRatio    =   27
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
         Caption         =   "後視点"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "器械点"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "編集やり直し"
      Height          =   495
      Index           =   0
      Left            =   6600
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   495
      Index           =   1
      Left            =   6600
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Index           =   1
      Left            =   6600
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmSyokiset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InitSave()
    InitDT.Tensu = xTextN1(0).Text
    InitDT.x0 = xTextN1(1).Text
    InitDT.y0 = xTextN1(2).Text
    InitDT.z0 = xTextN1(3).Text
    InitDT.MH = xTextN1(4).Text
    InitDT.x1 = xTextN1(5).Text
    InitDT.y1 = xTextN1(6).Text
    InitDT.z1 = xTextN1(7).Text
    InitDT.AZIMUTH = xTextN1(8).Text
    
    Call WriteIni("計測条件", "測点数", CStr(InitDT.Tensu), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "器械点X", CStr(InitDT.x0), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "器械点Y", CStr(InitDT.y0), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "器械点Z", CStr(InitDT.z0), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "器械点MH", CStr(InitDT.MH), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "後視点X", CStr(InitDT.x1), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "後視点Y", CStr(InitDT.y1), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "後視点Z", CStr(InitDT.z1), CurrentDIR & "計測設定.ini")
    Call WriteIni("計測条件", "方向角", CStr(InitDT.AZIMUTH), CurrentDIR & "計測設定.ini")
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then Call InitSave
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    If Index = 0 Then
        xTextN1(0).Text = InitDT.Tensu
        xTextN1(1).Text = InitDT.x0
        xTextN1(2).Text = InitDT.y0
        xTextN1(3).Text = InitDT.z0
        xTextN1(4).Text = InitDT.MH
        xTextN1(5).Text = InitDT.x1
        xTextN1(6).Text = InitDT.y1
        xTextN1(7).Text = InitDT.z1
        xTextN1(8).Text = InitDT.AZIMUTH
        Command2(0).Enabled = False: Me.Refresh
        SyokiSW = False
    Else
        Call InitSave
        Command2(0).Enabled = False: Me.Refresh
        SyokiSW = False
    End If
    
End Sub

Private Sub Command4_Click()
    If xTextN1(0).Text < 1 Then
        MsgBox "測点数が設定されていません。", vbCritical
        xTextN1(1).SetFocus
    Else
        InitDT.Tensu = xTextN1(0).Text
        InitDT.x0 = xTextN1(1).Text
        InitDT.y0 = xTextN1(2).Text
        InitDT.z0 = xTextN1(3).Text
        InitDT.MH = xTextN1(4).Text
        InitDT.x1 = xTextN1(5).Text
        InitDT.y1 = xTextN1(6).Text
        InitDT.z1 = xTextN1(7).Text
        frmPOset.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    Top = 0
    Left = 0
    xTextN1(0).Text = InitDT.Tensu
    xTextN1(1).Text = InitDT.x0
    xTextN1(2).Text = InitDT.y0
    xTextN1(3).Text = InitDT.z0
    xTextN1(4).Text = InitDT.MH
    xTextN1(5).Text = InitDT.x1
    xTextN1(6).Text = InitDT.y1
    xTextN1(7).Text = InitDT.z1
    xTextN1(8).Text = InitDT.AZIMUTH
    SyokiSW = False
    Command2(0).Enabled = False: Me.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set RsctlFrm = Nothing
    End
    
End Sub

Private Sub xText1_ChangeText(Index As Integer)
    If SyokiSW = False Then
        SyokiSW = True
        Command2(0).Enabled = True: Me.Refresh
    End If
End Sub

Private Sub xText1_LostFocus(Index As Integer)
    If IsNumeric(xTextN1(Index).Text) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        xTextN1(Index).SetFocus
    End If
End Sub


