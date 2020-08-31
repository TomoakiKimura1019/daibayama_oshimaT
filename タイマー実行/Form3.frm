VERSION 5.00
Object = "{3C938673-0ADC-4ABD-94B8-BE648EC2C95C}#4.1#0"; "xNumN.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'å≈íË¬∞Ÿ ≥®›ƒﬁ≥
   Caption         =   "Form3"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3585
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.CommandButton Command2 
      Caption         =   "ÉLÉÉÉìÉZÉã"
      Height          =   735
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "åàíË"
      Height          =   735
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin xNumNLib.xNumN xNumN1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   615
      _Version        =   262145
      _ExtentX        =   1085
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      FormatNo        =   1
      IntFigs         =   2
      MaxValue        =   99
      MinValue        =   0
      Text            =   " 0"
      Caption         =   "xNumN1"
      CaptionRatio    =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      InputAreaBorder =   1
      TreatEnterAsTab =   -1  'True
      LostFocusAtFullInput=   -1  'True
      CancelKey       =   1
      ClearKey        =   3
      PMenuCaption0   =   "å≥Ç…ñﬂÇ∑(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "êÿÇËéÊÇË(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "∫Àﬂ∞(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "ì\ÇËïtÇØ(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "çÌèú(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "Ç∑Ç◊ÇƒëIë(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "èëéÆÉRÉsÅ[(&F)"
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
   Begin xNumNLib.xNumN xNumN1 
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   495
      _Version        =   262145
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      FormatNo        =   1
      IntFigs         =   2
      MaxValue        =   59
      MinValue        =   0
      ZeroFill        =   -1  'True
      Text            =   "00"
      Caption         =   "xNumN1"
      CaptionRatio    =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      InputAreaBorder =   1
      LostFocusAtFullInput=   -1  'True
      CancelKey       =   1
      ClearKey        =   3
      PMenuCaption0   =   "å≥Ç…ñﬂÇ∑(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "êÿÇËéÊÇË(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "∫Àﬂ∞(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "ì\ÇËïtÇØ(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "çÌèú(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "Ç∑Ç◊ÇƒëIë(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "èëéÆÉRÉsÅ[(&F)"
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
   Begin xNumNLib.xNumN xNumN1 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   495
      _Version        =   262145
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      FormatNo        =   1
      IntFigs         =   2
      MaxValue        =   59
      MinValue        =   0
      ZeroFill        =   -1  'True
      Text            =   "00"
      Caption         =   "xNumN1"
      CaptionRatio    =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      InputAreaBorder =   1
      LostFocusAtFullInput=   -1  'True
      CancelKey       =   1
      ClearKey        =   3
      PMenuCaption0   =   "å≥Ç…ñﬂÇ∑(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "êÿÇËéÊÇË(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "∫Àﬂ∞(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "ì\ÇËïtÇØ(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "çÌèú(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "Ç∑Ç◊ÇƒëIë(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "èëéÆÉRÉsÅ[(&F)"
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
      Caption         =   "éû      ï™     ïb"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hr As Integer
Dim mn As Integer
Dim se As Integer

Dim tIntv As Date

Private Sub Command1_Click()
    If ivCode = 0 Then
        KE_intv = TimeSerial(xNumN1(0).Value, xNumN1(1).Value, xNumN1(2).Value)
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    If ivCode = 0 Then
        tIntv = KE_intv
    End If

    Caption = "ÉCÉìÉ^Å[ÉoÉãê›íË"
    If tIntv < 1 Then
        hr = Hour(tIntv)
    Else
        hr = Fix(tIntv * 24)
    End If
    mn = Minute(tIntv)
    se = Second(tIntv)
    
    xNumN1(0).Value = hr
    xNumN1(1).Value = mn
    xNumN1(2).Value = se
    
End Sub
