VERSION 5.00
Object = "{C1C35723-CBFC-101C-924B-7652655053A6}#4.2#0"; "xdate.ocx"
Object = "{C1C3572D-CBFC-101C-924B-7652655053A6}#4.2#0"; "xtime.ocx"
Begin VB.Form frmIntvNew 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�C���^�[�o���ݒ�"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   ClipControls    =   0   'False
   Icon            =   "frmIntvNew.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   6690
   Begin VB.CommandButton Command1 
      Caption         =   "�ۑ�"
      Height          =   495
      Index           =   1
      Left            =   5160
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���j���[�ɖ߂�"
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin xTimeLib.xTime xTime2 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   735
      _Version        =   262146
      _ExtentX        =   1305
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.74
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
      HourValue       =   0
      MinuteValue     =   0
      SecondValue     =   0
      DisplayString   =   "00:00"
      SpinButton      =   0   'False
      FocusBorder     =   1
      SecondInput     =   0   'False
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      InputAreaBorder =   0
      TreatEnterAsTab =   -1  'True
      PMenuCaption0   =   "���ɖ߂�(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "�؂���(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "��߰(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "�\��t��(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "�폜(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "���ׂđI��(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "�����R�s�[(&F)"
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
   Begin xDateLib.xDate xDate1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   1275
      _Version        =   262146
      _ExtentX        =   2249
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.74
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignmentH      =   1
      BorderStyle     =   4
      DateFormat      =   3
      Value           =   36871
      YearValue       =   2000
      MonthValue      =   12
      DayValue        =   11
      Week            =   2
      DisplayString   =   "2000/12/11"
      SpinButton      =   0   'False
      FocusBorder     =   1
      MaxDate         =   "2030/12/31"
      ZeroFill        =   -1  'True
      Caption         =   "�ϑ���"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      InputAreaBorder =   0
      TreatEnterAsTab =   -1  'True
      PMenuCaption0   =   "���ɖ߂�(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "�؂���(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "��߰(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "�\��t��(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "�폜(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "���ׂđI��(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "�����R�s�[(&F)"
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
   Begin xTimeLib.xTime xTime1 
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   1050
      _Version        =   262146
      _ExtentX        =   1852
      _ExtentY        =   661
      _StockProps     =   79
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.74
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlignmentH      =   1
      BorderStyle     =   4
      TimeFormat      =   4
      HourValue       =   0
      MinuteValue     =   0
      SecondValue     =   0
      DisplayString   =   "00:00:00"
      SpinButton      =   0   'False
      FocusBorder     =   1
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.99
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionRatio    =   0
      InputAreaBorder =   0
      TreatEnterAsTab =   -1  'True
      PMenuCaption0   =   "���ɖ߂�(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   0   'False
      PSeparator0     =   0   'False
      PMenuCaption1   =   "�؂���(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   0   'False
      PSeparator1     =   0   'False
      PMenuCaption2   =   "��߰(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   0   'False
      PSeparator2     =   0   'False
      PMenuCaption3   =   "�\��t��(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   0   'False
      PSeparator3     =   0   'False
      PMenuCaption4   =   "�폜(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   0   'False
      PSeparator4     =   0   'False
      PMenuCaption5   =   "���ׂđI��(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   0   'False
      PSeparator5     =   0   'False
      PMenuCaption6   =   "�����R�s�[(&F)"
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
   Begin VB.Label Label2 
      Caption         =   " ���v��������ِݒ聄 "
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Left            =   240
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "�C���^�[�o������"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  '�E����
      Caption         =   "����v�����t"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmIntvNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IntvCng As Boolean

Sub IntvSave(ck As Boolean)
    ck = True
    
    If xTime1.Value = 0 Then
        MsgBox "�C���^�[�o�����Ԃ́A0�b���傫�����Ԃ���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
        ck = False
        Exit Sub
    End If
    
    Keisoku_Time = xDate1.Value + xTime2.Value
    KE_intv = xTime1.Value
    Call IntvWrite
    
    �v��Form.xText2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    �v��Form.xText2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
    '�v�����ԍĐݒ�
    Call Ktime_ck
    
    IntvCng = False
End Sub
Private Sub Command1_Click(index As Integer)
    Dim ck As Boolean
    If index = 0 Then
        Unload Me
    Else
        Call IntvSave(ck)
        If ck = True Then MsgBox "�ۑ����������܂����B", vbInformation
    End If
End Sub

Private Sub Form_Load()
    
    Left = (Screen.Width - Me.Width) / 2
    top = 0
    
    xTime1.Value = KE_intv
    xDate1.Value = Format$(Keisoku_Time, "yyyy/mm/dd")
    xTime2.Value = Format$(Keisoku_Time, "hh:mm")
    
    xDate1.MinDate = Format$(Now, "yyyy/mm/dd")
    
    IntvCng = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As Integer, ck As Boolean
    
    If IntvCng = True Then
        Response = MsgBox("�ύX���ۑ�����Ă��܂���B�ۑ����܂����H", vbYesNoCancel + vbExclamation, "�I���̊m�F")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then
            Call IntvSave(ck)
            If ck = False Then Cancel = True: Exit Sub
        End If
    End If
    frmMenu.Show
End Sub

Private Sub xDate1_ChangeValue()
    IntvCng = True
End Sub

Private Sub xTime1_ChangeValue()
    IntvCng = True
End Sub

Private Sub xTime2_ChangeValue()
    IntvCng = True
End Sub
