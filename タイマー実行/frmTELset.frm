VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmTELset 
   Caption         =   "緊急時連絡先入力"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14880
   Icon            =   "frmTELset.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Frame Frame2 
      Caption         =   "〔 説明 〕"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   12975
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   12735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Index           =   0
      Left            =   13440
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   495
      Index           =   1
      Left            =   13440
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "セル編集"
      Height          =   2175
      Left            =   13320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "コピー"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "切り取り"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "貼り付け"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
   End
   Begin VCF150Ctl.F1Book F1Book1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5530
      _0              =   $"frmTELset.frx":08CA
      _1              =   $"frmTELset.frx":0CD3
      _2              =   $"frmTELset.frx":10DC
      _3              =   $"frmTELset.frx":14E5
      _4              =   $"frmTELset.frx":18EE
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "frmTELset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw As Boolean
Dim maxFLD As Integer
Dim L(500) As String, co As Integer
Dim DM(500) As String, DMco As Integer
Dim Tabl_path As String
Public CurrentDIR As String   'keisoku.exeがあるフォルダ

Sub InitSheet()
    Dim f As Integer, L As String
    Dim i As Integer, j As Integer
    Dim Dco As Long, Dsw As Boolean
    Dim maxREC As Long, Rco As Long, Rpst1 As Long, Rpst2 As Long, Rsw As Boolean
    Dim FLDno As Integer
    Dim FILENAME As String

    On Error Resume Next
    Screen.MousePointer = 11
    
    F1Book1.DoSafeEvents = True
    F1Book1.EnableProtection = True
    
    F1Book1.MaxCol = 7
    F1Book1.ColText(1) = "会社名１"
    F1Book1.ColText(2) = "会社名２"
    F1Book1.ColText(3) = "会社名３"
    F1Book1.ColText(4) = "電話番号１"
    F1Book1.ColText(5) = "電話番号２"
    F1Book1.ColText(6) = "担当者正"
    F1Book1.ColText(7) = "担当者副"
    
    FILENAME = Tabl_path & "緊急連絡.dat"
    If Dir(FILENAME) <> "" Then
        f = FreeFile
        Open FILENAME For Input Shared As #f
            i = 0: DMco = 0
            Do While Not (EOF(f))
                Line Input #f, L
                If Left$(L, 1) = ":" Then
                    DMco = DMco + 1
                    DM(DMco) = L
                Else
                    i = i + 1
                    F1Book1.TextRC(i, 1) = Trim(SEEKmoji(L, 5, 30))
                    F1Book1.TextRC(i, 2) = Trim(SEEKmoji(L, 35, 30))
                    F1Book1.TextRC(i, 3) = Trim(SEEKmoji(L, 65, 30))
                    F1Book1.TextRC(i, 4) = Trim(SEEKmoji(L, 95, 20))
                    F1Book1.TextRC(i, 5) = Trim(SEEKmoji(L, 115, 20))
                    F1Book1.TextRC(i, 6) = Trim(SEEKmoji(L, 135, 16))
                    F1Book1.TextRC(i, 7) = Trim(SEEKmoji(L, 151, 16))
                    
                End If
            Loop
        Close #f
    End If

    F1Book1.SetSelection 1, 1, F1Book1.MaxRow, F1Book1.MaxCol
    F1Book1.SetProtection False, True
    F1Book1.EnableProtection = True
    F1Book1.SetSelection i + 1, 1, i + 1, 1
    F1Book1.ShowActiveCell
    
    F1Book1.Modified = False
    F1Book1.DoSafeEvents = True
    F1Book1.SetFocus
    
    Screen.MousePointer = 0
End Sub

Sub SaveData()
    Dim i As Integer, f As Integer
    Dim j As Integer
    Dim SS As String
    Dim Incel As Integer
    
    f = FreeFile
    Open Tabl_path & "緊急連絡.dat" For Output Lock Write As #f
        For i = 1 To DMco
            Print #f, DM(i)
        Next i
        
        For i = 1 To F1Book1.MaxRow
            SS = F1Book1.TextRC(i, 1) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 1), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 2) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 2), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 3) & Space$(30 - LenB(StrConv(F1Book1.TextRC(i, 3), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 4) & Space$(20 - LenB(StrConv(F1Book1.TextRC(i, 4), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 5) & Space$(20 - LenB(StrConv(F1Book1.TextRC(i, 5), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 6) & Space$(16 - LenB(StrConv(F1Book1.TextRC(i, 6), vbFromUnicode)))
            SS = SS & F1Book1.TextRC(i, 7) & Space$(16 - LenB(StrConv(F1Book1.TextRC(i, 7), vbFromUnicode)))
            If Trim(SS) <> "" Then Print #f, Format(i, "!@@@@") & SS
        Next i
    Close #f
    
    F1Book1.Modified = False
End Sub

Private Sub Command1_Click(Index As Integer)
'    On Error Resume Next
    
    If Index = 0 Then
        Call SaveData
        MsgBox "保存が完了しました。", vbInformation
        F1Book1.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub Command3_Click(Index As Integer)
On Error Resume Next
    
    If Index = 0 Then F1Book1.EditCopy: F1Book1.EditClear F1ClearValues '切り取り
    If Index = 1 Then F1Book1.EditCopy        'コピー
    If Index = 2 Then F1Book1.EditPasteValues '貼り付け
    F1Book1.SetFocus
End Sub

Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
'    On Error Resume Next
'    Select Case F1Book1.Col
'    Case 1
'        If IsDate(EditString) = True Then
'            F1Book1.TextRC(F1Book1.Row, 1) = Format(DateValue(EditString), "yyyy/mm/dd")
'            F1Book1.CancelEdit
'            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
'        Else
'            MsgBox "日付と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
'            F1Book1.CancelEdit
'        End If
'    Case 2
'        If IsDate(EditString) = True Then
'            F1Book1.TextRC(F1Book1.Row, 2) = Format(TimeValue(EditString), "h:nn")
'            F1Book1.CancelEdit
'            F1Book1.SetActiveCell F1Book1.Row + 1, F1Book1.Col
'        Else
'            MsgBox "時間と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
'            F1Book1.CancelEdit
'        End If
'    Case Else
'        If IsNumeric(EditString) = False Then
'            MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
'            F1Book1.CancelEdit
'        End If
'        If F1Book1.TextRC(F1Book1.Row, 1) = "" Or F1Book1.TextRC(F1Book1.Row, 2) = "" Then
'            MsgBox "計測日付・計測時間が入力されていません。先に、計測日付・計測時間を入力してください。", vbCritical, "エラーメッセージ"
'            F1Book1.CancelEdit
'        End If
'    End Select
End Sub

Private Sub F1Book1_SelChange()
    Dim SS As String
        
    Select Case F1Book1.Col
    Case 1 To 3
        SS = "会社名を入力します｡最大文字数は､全角文字で15文字までです｡"
    Case 4 To 5
        SS = "電話番号を入力します｡最大文字数は､全角文字で10文字までです｡"
    Case Else
        SS = "担当者を入力します｡最大文字数は､全角文字で8文字までです｡"
    End Select
    Label1.Caption = SS
End Sub

Private Sub Form_Load()
    ChDrive App.Path
    ChDir App.Path
    
    CurrentDIR = App.Path
    If Right(CurrentDIR, 1) = "\" Then Else CurrentDIR = CurrentDIR & "\"
    
    Tabl_path = GetIni("フォルダ名", "環境データ", CurrentDIR & "連絡設定.ini")
    Call InitSheet
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Response As Integer
    If F1Book1.Modified Then
        Response = MsgBox("変更が保存されていません。保存しますか？", vbYesNoCancel + vbExclamation, "終了の確認")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then Call SaveData
    End If
End Sub
'**********************************************************************************************
'   フォームのサイズを変更した場合にコントロールの位置を設定
'**********************************************************************************************
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Frame1
        .Left = ScaleWidth - 1560
        .Top = 120
    End With
    With Frame2
        .Left = 120
        .Top = ScaleHeight - 1125
    End With
    
    With Command1(0)
        .Left = ScaleWidth - 1440
        .Top = 2760
    End With
    With Command1(1)
        .Left = ScaleWidth - 1440
        .Top = ScaleHeight - 645
    End With
    
    With F1Book1
        .Left = 120
        .Top = 120
        .Height = Me.ScaleHeight - 1350
        .Width = Me.ScaleWidth - 1905
    End With
'Debug.Print Me.Width, Me.Height
End Sub


'**********************************************************************************************
'   文字列から指定した文字数分の文字列を返します。
'   （全角文字を２文字、半角文字を１文字とします。）
'**********************************************************************************************
Public Function SEEKmoji(strCheckString As String, mojiST As Integer, mojiMAX As Integer) As String

    'Forカウンタ
    Dim i As Long
    
    '調査対象文字列の長さを格納
    Dim lngCheckSize As Long
    
    'ANSIへの変換後の文字を格納
    Dim lngANSIStr As Long
    
    Dim co As Integer '文字数
    Dim SS As String
    
    lngCheckSize = Len(strCheckString)

    co = 0: SS = ""
    For i = 1 To lngCheckSize
        'StrConvでUnicodeからANSIへと変換
        lngANSIStr = LenB(StrConv(Mid(strCheckString, i, 1), vbFromUnicode))
        
        co = co + lngANSIStr
        If co >= mojiST And co < (mojiST + mojiMAX) Then
            SS = SS + Mid(strCheckString, i, 1)
        End If
    Next i
    SEEKmoji = SS
End Function

