VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmSetTABLE 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "環境設定"
   ClientHeight    =   9570
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "SetTABLE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   8475
   StartUpPosition =   2  '画面の中央
   Begin VCF150Ctl.F1Book F1Book1 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   16510
      _0              =   $"SetTABLE.frx":0442
      _1              =   $"SetTABLE.frx":084B
      _2              =   $"SetTABLE.frx":0C54
      _3              =   $"SetTABLE.frx":105D
      _4              =   $"SetTABLE.frx":1466
      _5              =   $"SetTABLE.frx":186F
      _6              =   $"SetTABLE.frx":1C78
      _7              =   $"SetTABLE.frx":2081
      _8              =   $"SetTABLE.frx":248A
      _count          =   9
      _ver            =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   20000
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   20000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ"
      Begin VB.Menu mnuSave 
         Caption         =   "保存"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "印刷"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "終了"
      End
   End
End
Attribute VB_Name = "frmSetTABLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim L(500) As String, co As Integer
'Public OWARI As Boolean

Private Sub FileSave(SW As Boolean)
    Dim i As Integer, f As Integer
    Dim j As Integer
    
    Dim Incel As Integer
    
    SW = False
    For i = 1 To F1Book1.MaxRow
        If F1Book1.TextRC(i, 4) = "" Then SW = True: Incel = 4: Exit For
        If F1Book1.TextRC(i, 5) = "" Then SW = True: Incel = 5: Exit For
        If F1Book1.TextRC(i, 6) = "" Then SW = True: Incel = 6: Exit For
        If F1Book1.TextRC(i, 7) = "" Then SW = True: Incel = 7: Exit For
    Next i
    If SW = True Then
        F1Book1.SetActiveCell i, Incel
        MsgBox "空白のセルが見つかりました。必ず数値を入力してください。", vbCritical, "エラーメッセージ"
        Exit Sub
    End If
    
    f = FreeFile
    Open KEISOKU.Tabl_path & CTABLE1_DAT For Output Lock Write As #f
        j = 0
        For i = 1 To co
            If Left$(L(i), 1) = ":" Then
                Print #f, L(i)
            Else
                j = j + 1
                Print #f, Left$(L(i), 16);
                Print #f, Format$(F1Book1.TextRC(j, 4), "@@@@");
                Print #f, Format$(F1Book1.TextRC(j, 5), "@@@@@@@@");
                Print #f, Format$(F1Book1.TextRC(j, 6), "@@@@@@@@@@");
                
                If F1Book1.TextRC(j, 7) = "******" Then
                    Print #f, " -999999";
                Else
                    Print #f, Format$(F1Book1.TextRC(j, 7), "@@@@@@@@");
                End If
                If F1Book1.TextRC(j, 8) = "******" Then
                    Print #f, " -999999"
                Else
                    Print #f, Space$(8 - LenB(StrConv(F1Book1.TextRC(j, 8), vbFromUnicode))) & F1Book1.TextRC(j, 8)
                End If
            End If
        Next i
    Close #f
    F1Book1.Modified = False
End Sub


Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
    If F1Book1.Col = 8 Then Exit Sub
    
    On Error Resume Next
    If IsNumeric(EditString) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        F1Book1.CancelEdit
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim f As Integer
    Dim Dpo As Integer
    Dim SS As String
    Dim dan_ID As Integer, kou_ID As Integer, ten_ID As Integer
    Dim oldKOU As Integer
    
    Dim lBorder As Integer
    Dim rBorder As Integer
    Dim tBorder As Integer
    Dim bBorder As Integer
    Dim shade As Integer
    Dim lColor As Long
    Dim rColor As Long
    Dim tColor As Long
    Dim bColor As Long
    
    frmCLOSE.setTABLE = False
    Width = 8565
    Height = 10230
    
    i = FileCheck(KEISOKU.Tabl_path & CTABLE1_DAT, "環境データ")
    If i = 0 Then Unload 計測Form: End

    co = 0: Dpo = 0
    f = FreeFile
    Open KEISOKU.Tabl_path & CTABLE1_DAT For Input Shared As #f
        Do While Not (EOF(f))
            co = co + 1
            Line Input #f, L(co)
            If Left$(L(co), 1) <> ":" Then
                Dpo = Dpo + 1
                kou_ID = CInt(Mid$(L(co), 1, 4))
                dan_ID = CInt(Mid$(L(co), 5, 4))
                ten_ID = CInt(Mid$(L(co), 9, 4))
                
                F1Book1.TextRC(Dpo, 1) = Trim$(DanSet(kou_ID, dan_ID).ti)
                
                SS = Trim$(kou(kou_ID, 1).ti1)
                F1Book1.TextRC(Dpo, 2) = SS
                F1Book1.TextRC(Dpo, 3) = ten_ID
                
                F1Book1.TextRC(Dpo, 4) = Trim$(Mid$(L(co), 17, 4))
                F1Book1.TextRC(Dpo, 5) = Trim$(Mid$(L(co), 21, 8))
                F1Book1.TextRC(Dpo, 6) = Trim$(Mid$(L(co), 29, 10))
                
                If Trim$(Mid$(L(co), 39, 8)) = "-999999" Then
                    F1Book1.TextRC(Dpo, 7) = "******"
                    F1Book1.SetActiveCell Dpo, 7
                    F1Book1.SetProtection True, True
                Else
                    F1Book1.TextRC(Dpo, 7) = Trim$(Mid$(L(co), 39, 8))
                End If
                If Trim$(Mid$(L(co), 47, 8)) = "-999999" Then
                    F1Book1.TextRC(Dpo, 8) = "******"
                    F1Book1.SetActiveCell Dpo, 8
                    F1Book1.SetProtection True, True
                Else
                    F1Book1.TextRC(Dpo, 8) = Trim$(SEEKmoji(L(co), 47, 8))
                End If
                
                If ten_ID = 1 And Dpo > 1 Then
                    F1Book1.SetSelection Dpo, 1, Dpo, F1Book1.MaxCol
                    
                    F1Book1.GetBorder lBorder, rBorder, tBorder, bBorder, shade, lColor, rColor, tColor, bColor
                    F1Book1.SetBorder -1, lBorder, rBorder, 1, bBorder, shade, -1, lColor, rColor, rColor, bColor
                End If
                oldKOU = kou_ID
            End If
        Loop
    Close #f
    
    F1Book1.MaxRow = Dpo
    F1Book1.EnableProtection = True
    F1Book1.SetActiveCell 1, 4
    F1Book1.DoSafeEvents = True
    F1Book1.Modified = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As Integer
    Dim SW As Boolean
    
    If F1Book1.Modified Then
        Response = MsgBox("変更が保存されていません。保存しますか？", vbYesNoCancel + vbExclamation, "終了の確認")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then
            Call FileSave(SW)
            If SW = True Then Cancel = True: Exit Sub
        End If
    End If
    frmCLOSE.setTABLE = True
End Sub

Private Sub mnuEnd_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo trap1
    
    F1Book1.SetSelection 1, 1, F1Book1.MaxRow, F1Book1.MaxCol
    F1Book1.SetPrintAreaFromSelection
    F1Book1.FilePrintEx True, False
    F1Book1.SetActiveCell 1, 4
    Me.Refresh
trap1:
End Sub

Private Sub mnuSave_Click()
    Dim SW As Boolean
    
    If F1Book1.Modified = True Then Call FileSave(SW)
    If SW = False Then MsgBox "保存が終了しました。", vbInformation
End Sub

