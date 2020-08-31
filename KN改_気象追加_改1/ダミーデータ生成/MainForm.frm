VERSION 5.00
Object = "{7F11ED83-882D-4ED8-A1B2-E149DE36F7B0}#4.1#0"; "xTextN.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Form3"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3645
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
   ScaleHeight     =   3960
   ScaleWidth      =   3645
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   3000
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Left            =   2400
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   1800
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   8400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3201
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox List1 
      Height          =   1140
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CheckBox Check2 
      Caption         =   "有効"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   600
      Value           =   1  'ﾁｪｯｸ
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "有効"
      Height          =   340
      Left            =   8520
      TabIndex        =   20
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      TabIndex        =   19
      ToolTipText     =   "パラメーターをリロード"
      Top             =   120
      Width           =   255
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3660
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6376
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
      Left            =   9720
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
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
      Left            =   9720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
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
      Left            =   9720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
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
   Begin xTextNLib.xTextN xTextN3 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
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
      InputAreaColor  =   16777152
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
   Begin xTextNLib.xTextN xTextN3 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   840
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
      InputAreaColor  =   16777152
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
   Begin xTextNLib.xTextN xTextN3 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
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
      InputAreaColor  =   16777152
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
      Left            =   1560
      TabIndex        =   16
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
   Begin VB.Label Label11 
      Caption         =   "動作ログ"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label10 
      Height          =   255
      Left            =   9840
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  '右揃え
      Caption         =   "現在時刻"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   180
      Width           =   855
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   120
      Top             =   720
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   8400
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "インターバル"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1290
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "前回記録時間"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "次回記録時間"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "インターバル"
      Height          =   255
      Left            =   8520
      TabIndex        =   9
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "前回測定時間"
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "次回測定時間"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      ToolTipText     =   "ダブルクリックで設定変更"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  '右揃え
      Caption         =   "単位：m"
      Height          =   255
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "最新計測値"
      Height          =   255
      Left            =   8520
      TabIndex        =   2
      Top             =   2520
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

'===================
'ＴＤＳ読み込みデータフォーマット（1回分）
Private Type TDSdt
    Mdate As String * 19     '計測日時
    tDATA(0 To 99) As dt '測定値
End Type

Dim STtime As Date, minTIME As Date
Dim keisoku_f As Boolean
Dim kiroku_f As Boolean
Dim Thistime As String
Dim Stat As Integer

Const TDSDATALEN As Integer = 8  '530の場合
                                 '303だと8

Dim RSlog As Integer         '通信ログを残すかどうか
Const fRSlog As Integer = 9  '通信ログ用ファイル番号

Const TDSfm As Integer = 6
'TDSfm = 4      '303
'TDSfm = 6      '150
    
Dim seigen() As Double, tCH() As Integer

Dim ButsuriRyou(10) As Double

Dim fZorder As Integer

Private Sub Command1_Click()
    Call DataPrint
End Sub

Private Sub Form_Resize()
    Dim TMP As String
    TMP = GetIni("Form", "Zorder", App.Path & "\計測設定.ini")
    
    If TMP = "-1" Then
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

Private Sub Form_Load()

    '元の表示位置に表示するための設定
    Dim fTop As Long, fLeft As Long
    Dim TMP0 As Variant
    TMP0 = GetIni("Form", "top", CurrentDir & "計測設定.ini")
    If TMP0 <> "" Then
        fTop = TMP0
        If fTop < 0 Then fTop = 0
    End If
    TMP0 = GetIni("Form", "left", CurrentDir & "計測設定.ini")
    If TMP0 <> "" Then
        fLeft = TMP0
        If fLeft < 0 Then fLeft = 0
    End If
    
    Top = fTop
    Left = fLeft
    
    List1.Clear
    
    '起動ログ
    Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 起動")
    
    Caption = "SKダミー"
    
    Me.Enabled = False
    
    Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
    Call nextDate(kiroku_TimeZ, kiroku_intv, kiroku_Time)
    '次回計測時間計算
'    Debug.Print toTMSstring(KE_intv)
    Call IntvWrite
    xTextN2(0).Text = Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN2(1).Text = toTMSstring(KE_intv)
    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
    xTextN3(0).Text = Format$(kiroku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN3(1).Text = toTMSstring(kiroku_intv)
    xTextN3(2).Text = Format$(kiroku_Time, "yyyy/mm/dd hh:nn:ss")
    
    ListView1.FullRowSelect = True
    ListView1.GridLines = True
    Call ListView1.ColumnHeaders.Add(, , "CH", 600)
    Call ListView1.ColumnHeaders.Add(, , "測定値", 1200, lvwColumnRight)
'    Call ListView1.ColumnHeaders.Add(, , "補正値", 1200, lvwColumnRight)
'    Call ListView1.ColumnHeaders.Add(, , "Y方向", 900, lvwColumnRight)
'    Call ListView1.ColumnHeaders.Add(, , "Z方向", 900, lvwColumnRight)

    Call ListView1.ListItems.Clear
    
    Dim ddat(4) As String
        
    Call fadd(ddat, "", 0)
    
    keihou_L = 0
    Call DataPrint
    
    Me.Enabled = True
    
    
'
'AddMaster  'test
'

    
    
    '自動計測開始
    keisoku_f = False
    kiroku_f = False
'    Me.SetFocus
    Timer1.Interval = 200
    Timer1.Enabled = True
    Timer2.Interval = 250
    Timer2.Enabled = True
    
'TeidenMail
End Sub

Private Sub fadd(DAT() As String, nm As String, ia As Integer)
    Dim obItem          As ListItem
    Set obItem = ListView1.ListItems.Add()
    obItem.Text = nm
'    obItem.SubItems(1) = Format(dat(1), "+0.0000;-0.0000")
'    obItem.SubItems(2) = Format(dat(2), "+0.0000;-0.0000")
    obItem.SubItems(1) = DAT(3) 'Format(dat(3), "+0.0000;-0.0000")
'    obItem.SubItems(2) = Format(dat(4), vFMT(ia))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer, ENDsw As Boolean
    
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
        
        Call WriteIni("Form", "top", (Top), CurrentDir & "計測設定.ini")
        Call WriteIni("Form", "left", (Left), CurrentDir & "計測設定.ini")
       
        If tcpClient.State = sckConnected Then tcpClient.Close

       '終了ログ
        Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 終了")
        
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
    Dim ret As Integer
    Dim pa As String
    Dim i As Integer
    Dim f As Integer
    
        mdy = Now
        Thistime = Format$(mdy, "yyyy/mm/dd hh:nn:ss")
        xTextNtime.Text = Format$(Thistime, "YYYY/MM/DD hh:mm:ss")
        
        If Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss") = Thistime Then keisoku_f = True
        If Format$(kiroku_Time, "yyyy/mm/dd hh:nn:ss") = Thistime Then kiroku_f = True
        
        If Check1.Value <> 1 Then keisoku_f = False
        If Check2.Value <> 1 Then kiroku_f = False
        
        If keisoku_f = True Or kiroku_f = True Then

            MainForm.StatusBar1.Panels(1).Text = "*** 計測中 ***"
            MainForm.Enabled = False
    
            pa = KNpath(1)
            If Right$(pa, 1) <> "\" Then pa = pa & "\"
            pa = pa & YMD2PathName(Thistime) '& "\R20171117080000_TOTAL.TXT"
            
            f = FreeFile
            Open pa For Output As #f
            For i = 1 To 14
                Print #f, Thistime;
                Print #f, ","; "O,"; i; ",0,0,0,-0.3,0,-5.5,0,-5.5,0,6.4,0,1,0,1.1,0,1.2,0,24.7537"
            Next i
            Close #f
            
            pa = KNpath(2)
            If Right$(pa, 1) <> "\" Then pa = pa & "\"
            pa = pa & YMD2PathName(Thistime) '& "\R20171117080000_TOTAL.TXT"
            
            f = FreeFile
            Open pa For Output As #f
            For i = 1 To 13
                Print #f, Thistime;
                Print #f, ","; "N,"; i; ",0,0,0,-0.3,0,-5.5,0,-5.5,0,6.4,0,1,0,1.1,0,1.2,0,24.7537"
            Next i
            For i = 1 To 9
                Print #f, Thistime;
                Print #f, ","; "T,"; i; ",0,0,0,-0.3,0,-5.5,0,-5.5,0,6.4,0,1,0,1.1,0,1.2,0,24.7537"
            Next i
            Close #f
            
            '次回計測時間計算
            If keisoku_f = True Then
                keisoku_f = False
                Keisoku_TimeZ = Keisoku_Time
                Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
            End If
            If kiroku_f = True Then
                kiroku_f = False
                kiroku_TimeZ = kiroku_Time
                Call nextDate(kiroku_TimeZ, kiroku_intv, kiroku_Time)
            End If
            Call IntvWrite

            Call DayTimeWrite
            
            MainForm.Enabled = True
            If Stat = 0 Then MainForm.StatusBar1.Panels(1).Text = ""
            
        End If
            
99      '計測時間がすぎた場合
        If 0 <= DateDiff("s", Keisoku_Time, mdy) Then 'If nt < Now Then
            Call nextDate(Keisoku_TimeZ, KE_intv, Keisoku_Time)
            Call IntvWrite
            
            xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            
            If Check1.Value = 1 Then
                'ログ
                Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 計測時間が過ぎていたため、再設定しました。")
            End If
        End If
        If 0 <= DateDiff("s", kiroku_Time, mdy) Then 'If nt < Now Then
            Call nextDate(kiroku_TimeZ, kiroku_intv, kiroku_Time)
            Call IntvWrite
            
            xTextN3(2).Text = Format$(kiroku_Time, "yyyy/mm/dd hh:nn:ss")
            
            'ログ
            If Check2.Value = 1 Then
                Call WriteEvents(Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : 記録時間が過ぎていたため、再設定しました。")
            End If
        End If

End Sub

Private Sub DayTimeWrite()
    xTextN2(0).Text = Format$(Keisoku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN2(1).Text = toTMSstring(KE_intv)
    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")

    xTextN3(0).Text = Format$(kiroku_TimeZ, "yyyy/mm/dd hh:nn:ss")
    xTextN3(1).Text = toTMSstring(kiroku_intv)
    xTextN3(2).Text = Format$(kiroku_Time, "yyyy/mm/dd hh:nn:ss")
End Sub

Public Sub DataPrint()
    Dim f As Integer
    Dim sbf As String
    Dim i As Integer
    Dim ten_ID As Integer
    Dim Udt As TDSdt
    Dim t_id As Integer
    Dim FLDno  As Integer
    Dim kou_ID As Integer, dan_ID As Integer ', t_ID As Integer
    
    Dim kDate As String
    Dim kCH(100) As String
    Dim kDATA(100) As String
    
    Dim DTL As Integer
    
    'ＴＤＳ変数初期化
    For i = 1 To TDS_CH: dt1(i) = 999999:  Next i
    Udt.tDATA(0).CH = 0
    
On Error Resume Next
    
    'ＴＤＳデータ読み込み
    sbf = ""
    If Dir$(CurrentDir & "final.dat") <> "" Then
        f = FreeFile
        Open CurrentDir & "final.dat" For Input Access Read Lock Write As #f
        Line Input #f, sbf
        If IsDate(sbf) = False Then
            Close #f
            Exit Sub
        End If
        
        kDate = sbf
        i = 0
        Do While Not (EOF(f))
            Line Input #f, sbf
            If 11 <= Len(sbf) Then
                If 0 < InStr(sbf, "E") Then Exit Do
                i = i + 1
                If InStr("0123456789", Left(sbf, 1)) = 0 Then
                    kCH(i) = Mid$(sbf, 2, 3)
                    kDATA(i) = Right$(sbf, Len(sbf) - 4)
                Else
                    kCH(i) = Mid$(sbf, 1, 3)
                    kDATA(i) = Right$(sbf, Len(sbf) - 3)
                End If
            End If
        Loop
        Close #f
        Label10.Caption = kDate

        DTL = i
        '物理量計算
        kou_ID = 1: dan_ID = 1: t_id = 1
            
        Dim ia As Integer
        Dim ddat(4) As String, nm As String
        Call ListView1.ListItems.Clear
        
        Dim tDATA As Double
    
        For ia = 1 To DTL
            nm = kCH(ia)
            ddat(3) = kDATA(ia)
            Call fadd(ddat, nm, ia)
        Next ia
    
    End If
    
On Error GoTo 0
    
    Exit Sub
    
    '    Dim DeltaT As Double '温度差℃
    '    For ten_ID = 1 To TDSTbl(0).ch
    '        ia = TDSTbl(ten_ID).FLD
    '        If OndoCH = TDSTbl(ten_ID).ch Then
    '            DeltaT = ((kDATA(ia) - TDSTbl(ia).syo) * TDSTbl(ia).kei) - stOndo
    '        End If
    '    Next ten_ID
        
        For ten_ID = 1 To TDSTbl(0).CH
            ia = TDSTbl(ten_ID).FLD
            If IsNumeric(kDATA(ia)) = True Then
                If TDSTbl(ia).kou = 1 Then
                    tDATA = (kDATA(ia) - TDSTbl(ia).Syo) * TDSTbl(ia).Kei '* 2.1 * 9.80665
                Else
                    tDATA = ((kDATA(ia) - TDSTbl(ia).Syo)) * TDSTbl(ia).Kei
                End If
                
                ddat(3) = kDATA(ia)
                ddat(4) = Format$(tDATA, vFMT(ia))
            Else
                ddat(3) = 999999
                ddat(4) = 999999
            End If
            nm = kCH(ia)
            Call fadd(ddat, nm, ia)
            ButsuriRyou(ia) = tDATA
        Next ten_ID
        
    'End If
End Sub

Private Function vFMT(No As Integer) As String
    If TDSTbl(No).keta = 0 Then
        vFMT = "@"
    Else
        vFMT = "." & String$(TDSTbl(No).keta, "0")
    End If
End Function

Private Sub AddMaster() '(ByVal kt As Date, TRGid As Integer)
    Dim f As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sa, sb As String
    Dim bf As String
    Dim hp As Integer
    
    ReDim DAT(tbl(0).CH) As String
    ReDim ddat(tbl(0).CH) As Double
    ReDim DAT(TDS_CH) As String
    ReDim ddat(TDS_CH) As Double
    For j = 0 To TDS_CH 'tbl(0).ch
        DAT(j) = "    999999"
    Next j

    f = FreeFile
    Open CurrentDir & "final.dat" For Input Access Read Lock Write As #f
    sb = StrConv(InputB$(LOF(1), 1), vbUnicode)
    Close #f

    sa = Split(sb, vbCrLf)

    bf = Mid$(sa(0), 1, 19)
    If 0 < InStr("0123456789", Left$(sa(1), 1)) Then
        hp = 1
    Else
        hp = 2
    End If

    For j = 1 To tbl(0).id
        For i = 1 To UBound(sa)
            If 0 < InStr(sa(i), "END") Then Exit For
            If 0 < InStr(sa(i), "/") Then
                'bf = Mid$(sA(i), 1, 19)
            Else
                If tbl(j).CH = Mid$(sa(i), hp, 3) Then
                    If 0 = InStr(sa(i), "--") And 0 = InStr(sa(i), "*") Then
                        DAT(g_cTable(j).id) = Right$(String$(10, " ") & Mid$(sa(i), hp + 3, Len(sa(i)) - (hp + 2)), 10)
                        Exit For
                    End If
                End If
            End If
        Next i
    Next j
    
'    bf = sa(0)
    For j = 1 To tbl(0).id
        ddat(j) = DAT(j)
        bf = bf & "," & (Right$(String$(10, " ") & ddat(j), 10))
    Next j


On Error Resume Next
    
    'master.datに歪みデータを保存 Append
    f = FreeFile
        Open KEISOKU.Data_path & "master.dat" For Append Lock Write As #f
        Print #f, bf
        Close #f

    '最深日時だけ保存
    f = FreeFile
    Open App.Path & "\newData.dat" For Output As #f
    Print #f, Left$(bf, 19)
    Close #f
    
        '送信用データの保存(TDS生値)
        'Dim sbf As String
        Dim pa As String
        Dim na As String
        Dim f1 As Integer ', f2 As Integer
        pa = GetIni("フォルダ名", "SendPath_S", CurrentDir & "計測設定.ini")
        If Right$(pa, 1) <> "\" Then
            pa = pa & "\"
        End If
        na = Format$(mdy, "YYYY-MM-DD_hh-nn-ss")
        na = na & ".csv"
        f1 = FreeFile
        Open pa & na For Output As #f1
        Print #f1, bf
        Close #f1

    If kiroku_f = False Then
        GoTo AM999
    End If
        
AM999:
On Error GoTo 0

End Sub

Private Function YMD2PathName(ss As String) As String
'R20171117080000_TOTAL.TXT
    Dim yy As String
    Dim mm As String
    yy = Format$(CDate(ss), "yyyymmddhhnnss")
    mm = Format$(CDate(ss), "mm")
    YMD2PathName = "R" & yy & "_TOTAL.TXT"
End Function
