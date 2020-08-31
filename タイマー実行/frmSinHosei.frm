VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmSinHosei 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "伸縮計 盛替え値入力"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   Icon            =   "frmSinHosei.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   6840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   1
      Top             =   8400
      Width           =   1095
   End
   Begin VCF150Ctl.F1Book F1Book1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   14631
      _0              =   $"frmSinHosei.frx":0442
      _1              =   $"frmSinHosei.frx":084C
      _2              =   $"frmSinHosei.frx":0C56
      _3              =   $"frmSinHosei.frx":105F
      _4              =   $"frmSinHosei.frx":1468
      _count          =   5
      _ver            =   2
   End
   Begin VB.Label Label1 
      Caption         =   "断面位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmSinHosei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L(3) As String

Dim kouNO As Integer, danNo As Integer, tenNO As Integer
Dim cmbNO As Integer, cmbDAN(20) As Integer, cmbTEN(20) As Integer
'Public OWARI As Boolean

Private Sub Combo1_Click()
    Dim f As Integer
    Dim i As Integer
    
    Dim LR As String
    Dim Response As Integer
    Dim FILENAME As String

On Error GoTo ERRSKIP

    If cmbNO = Combo1.ListIndex + 1 Then F1Book1.SetFocus: Exit Sub
    
    If cmbNO <> 0 And F1Book1.Modified Then
        Response = MsgBox("変更が保存されていません。保存しますか？", vbYesNo + vbExclamation, "確認")
        If Response = vbYes Then Call WriteSinHosei
    End If

    cmbNO = Combo1.ListIndex + 1
    F1Book1.ClearRange 1, 1, F1Book1.MaxRow, F1Book1.MaxCol, F1ClearValues
    
    FILENAME = KEISOKU.Tabl_path & "sin-" & CStr(cmbDAN(cmbNO)) & "-" & CStr(cmbTEN(cmbNO)) & ".Hos"
    If Dir$(FILENAME) <> "" Then
        f = FreeFile
        Open FILENAME For Input Shared As #f
            i = 0
            Do While Not (EOF(f))
                Line Input #f, LR
                i = i + 1
                F1Book1.TextRC(i, 1) = Format$(CDate(Mid$(LR, 1, 10)), "yyyy/mm/dd")
                F1Book1.TextRC(i, 2) = Format$(CDate(Mid$(LR, 11, 6)), "h:nn")
                F1Book1.TextRC(i, 3) = CSng(Mid$(LR, 17, 8))
            Loop
        Close #f
    
    End If
    
''    F1Book1.Row = 1: F1Book1.Col = 1: F1Book1.SetFocus
''    F1Book1.Modified = False

    F1Book1.SetSelection 1, 1, 30, 3
    F1Book1.SetProtection False, True
    F1Book1.EnableProtection = True
    F1Book1.SetSelection i + 1, 1, i + 1, 1
    F1Book1.ShowActiveCell
    F1Book1.Modified = False
    F1Book1.DoSafeEvents = True
    F1Book1.SetFocus
Exit Sub
ERRSKIP:
    Debug.Print Err.Number; Err.Description
    Resume Next
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub WriteSinHosei()
    Dim f As Integer, i As Integer
    Dim FILENAME As String
    
    FILENAME = KEISOKU.Tabl_path & "sin-" & CStr(cmbDAN(cmbNO)) & "-" & CStr(cmbTEN(cmbNO)) & ".Hos"
    f = FreeFile
    Open FILENAME For Output Lock Write As #f
        For i = 1 To 30
            If F1Book1.TypeRC(i, 1) = 0 Then Exit For
                
            Print #f, Format$(F1Book1.TextRC(i, 1), "@@@@@@@@@@");
            If F1Book1.TypeRC(i, 2) = 0 Then
                Print #f, "  0:00";
            Else
                Print #f, Format$(F1Book1.TextRC(i, 2), "@@@@@@");
            End If
            '''Print #f, Format$(F1Book1.TextRC(i, 3), "@@@@@@@@")
            Print #f, Space$(8 - LenB(StrConv(F1Book1.TextRC(i, 3), vbFromUnicode))) & F1Book1.TextRC(i, 3)
        Next i
    Close #f
    F1Book1.Modified = False
    
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then Call WriteSinHosei
    Unload Me
End Sub

Private Sub Command2_Click()
    Call WriteSinHosei
    MsgBox "保存が完了しました。", vbInformation
End Sub

Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
    On Error Resume Next
    
    Select Case F1Book1.Col
    Case 1
        If IsDate(EditString) = True Then
            F1Book1.TextRC(F1Book1.Row, 1) = Format$(DateValue(EditString), "yyyy/mm/dd")
            F1Book1.CancelEdit
            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
        Else
            MsgBox "日付と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            F1Book1.CancelEdit
        End If
    Case 2
        If IsDate(EditString) = True Then
            F1Book1.TextRC(F1Book1.Row, 2) = Format$(TimeValue(EditString), "h:nn")
            F1Book1.CancelEdit
            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
        Else
            MsgBox "時間と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            F1Book1.CancelEdit
        End If
    Case Else
        If IsNumeric(EditString) = False Then
            MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            F1Book1.CancelEdit
        End If
        If F1Book1.TextRC(F1Book1.Row, 1) = "" Or F1Book1.TextRC(F1Book1.Row, 2) = "" Then
            MsgBox "計測日付・計測時間が入力されていません。先に、計測日付・計測時間を入力してください。", vbCritical, "エラーメッセージ"
            F1Book1.CancelEdit
        End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer, co As Integer
    Dim SS As String
    
    Me.top = 計測Form.top + 800
    Me.Left = (Screen.Width - Width) / 2
    
'    F1Book1.DoSafeEvents = True
'    F1Book1.EnableProtection = True
    frmCLOSE.sinHosei = False
    
    co = 0: cmbNO = 0
    kouNO = 1
    danNo = 0
    For i = 1 To DanSet(kouNO, 0).dan
        For j = 1 To Tbl(kouNO, i, 0).ten
            co = co + 1
            cmbDAN(co) = i: cmbTEN(co) = j
            SS = Trim$(DanSet(kouNO, i).ti) & " " & Trim$(Tbl(kouNO, i, j).HAN)
            Combo1.AddItem Trim$(SS)
        Next j
    Next i
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Response As Integer
    
    If F1Book1.Modified Then
        Response = MsgBox("変更が保存されていません。保存しますか？", vbYesNoCancel + vbExclamation, "終了の確認")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then Call WriteSinHosei
    End If
    frmCLOSE.sinHosei = True
   
End Sub
