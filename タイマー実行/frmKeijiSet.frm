VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmKeijiSet 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "経時変化図設定"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   ForeColor       =   &H00000000&
   Icon            =   "frmKeijiSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   12150
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   375
      Index           =   1
      Left            =   10800
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   10800
      TabIndex        =   0
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "作図項目設定"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   10095
      Begin VB.CheckBox Check1 
         Caption         =   "表示"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "表示"
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   3210
         Index           =   3
         Left            =   7560
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   23
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   3
         Left            =   7560
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "表示"
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   20
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   3210
         Index           =   2
         Left            =   5160
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   19
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   2
         Left            =   5160
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "表示"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   3210
         Index           =   1
         Left            =   2760
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   15
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Height          =   3210
         Index           =   0
         Left            =   360
         Style           =   1  'ﾁｪｯｸﾎﾞｯｸｽ
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   7560
         TabIndex        =   22
         Text            =   "層別沈下計 累積変位"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5160
         TabIndex        =   18
         Text            =   "層別沈下計 累積変位"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Text            =   "層別沈下計 累積変位"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Text            =   "層別沈下計 累積変位"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   9840
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Caption         =   "No．４"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "No．３"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   5160
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "No．２"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "No．１"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         Height          =   4935
         Index           =   3
         Left            =   7440
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   4935
         Index           =   2
         Left            =   5040
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   4935
         Index           =   1
         Left            =   2640
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   4935
         Index           =   0
         Left            =   240
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ｙ軸設定"
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
      Height          =   1815
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   7095
      Begin VCF150Ctl.F1Book F1Book1 
         Height          =   1310
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2302
         _0              =   $"frmKeijiSet.frx":0442
         _1              =   $"frmKeijiSet.frx":084B
         _2              =   $"frmKeijiSet.frx":0C54
         _3              =   $"frmKeijiSet.frx":105D
         _4              =   $"frmKeijiSet.frx":1466
         _count          =   5
         _ver            =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ｘ軸設定"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text1 
         Alignment       =   1  '右揃え
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
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '右揃え
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
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "（最大時間＝24）"
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
         Left            =   2160
         TabIndex        =   30
         Top             =   435
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "表示時間数"
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
         Left            =   240
         TabIndex        =   6
         Top             =   440
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Top             =   920
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmKeijiSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kkou(Maxkou) As kkou1
Dim SelKou(Maxkou) As Integer
Dim SetKou(KeijiMAX, 3, 20) As Integer
Dim dan_ID As Integer, kou_ID As Integer
Dim stCK As Boolean
Dim MaxH As Integer
'Public OK As Boolean

Private Sub SokutenSet(no As Integer)
    Dim i As Integer, j As Integer
    
    For i = 1 To Tbl(Kkou(no + 1).kou, Kkou(no + 1).dan, 0).ten
        List1(no).AddItem Trim$(Tbl(Kkou(no + 1).kou, Kkou(no + 1).dan, i).HAN)
        For j = 1 To Kkou(no + 1).ten(0)
            If i = Kkou(no + 1).ten(j) Then List1(no).Selected(i - 1) = True: Exit For
        Next j
    Next i

End Sub

Private Sub Combo1_Click(Index As Integer)
    Dim i As Integer
    
'    If stCK = True And Kkou(Index + 1).kou = Combo1(Index).ListIndex + 1 Then Exit Sub
    
    Kkou(Index + 1).kou = SetKou(Index + 1, 1, Combo1(Index).ListIndex + 1)
    Kkou(Index + 1).type = SetKou(Index + 1, 2, Combo1(Index).ListIndex + 1)
    
    Combo2(Index).Clear
    If DanSet(Kkou(Index + 1).kou, 0).dan > 1 Then
        Combo2(Index).Enabled = True
        For i = 1 To DanSet(Kkou(Index + 1).kou, 0).dan
            Combo2(Index).AddItem Trim$(DanSet(Kkou(Index + 1).kou, i).ti)
        Next i
        If Kkou(Index + 1).ck = 1 Then Combo2(Index).ListIndex = Kkou(Index + 1).dan - 1 Else Combo2(Index).ListIndex = 0
    Else
        Combo2(Index).Enabled = False
        If DanSet(Kkou(Index + 1).kou, 0).dan = 1 Then Kkou(Index + 1).dan = 1
    End If
    List1(Index).Clear
    
    If Kkou(Index + 1).dan = 0 Then Exit Sub
    If Tbl(Kkou(Index + 1).kou, Kkou(Index + 1).dan, 0).ten = 0 Then Exit Sub
    
    Call SokutenSet(Index)

End Sub

Private Sub Combo2_Click(Index As Integer)
    
    If stCK = True And Kkou(Index + 1).dan = Combo2(Index).ListIndex + 1 Then Exit Sub
    
    Kkou(Index + 1).dan = Combo2(Index).ListIndex + 1
    
    If Kkou(Index + 1).ck = 1 And Combo2(Index).ListIndex = -1 Then Combo2(Index).ListIndex = 0
    If Kkou(Index + 1).ck = 0 Then Combo2(Index).ListIndex = 0
    
    
    List1(Index).Clear
    
    If Kkou(Index + 1).kou = 0 Then Exit Sub
    If Tbl(Kkou(Index + 1).kou, Kkou(Index + 1).dan, 0).ten = 0 Then Exit Sub
    
    Call SokutenSet(Index)
   
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer, j As Integer, co As Integer
    Dim ckSW As Boolean
    Dim k_ID As Integer
    Dim f As Integer
    Dim SS(3) As String
    
    If Index = 0 Then
        ckSW = True
        
        If IsNumeric(Text1(0).Text) = False Then
            MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            ckSW = False
            Exit Sub
        End If
        If Text1(0).Text <= 0 Then
            MsgBox "0より大きい数値を入力してください。", vbCritical, "エラーメッセージ"
            ckSW = False
            Exit Sub
        End If
        If IsNumeric(Text1(1).Text) = False Then
            MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
            ckSW = False
            Exit Sub
        End If
        If Text1(1).Text <= 0 Then
            MsgBox "0より大きい数値を入力してください。", vbCritical, "エラーメッセージ"
            ckSW = False
            Exit Sub
        End If
        
        If Text1(0).Text > MaxH Then
            MsgBox "表示時間数は、最大時間より小さい値を入力してください。", , "ﾒｯｾｰｼﾞ"
            ckSW = False
            Exit Sub
        End If
        For i = 1 To F1Book1.MaxRow

            If F1Book1.TextRC(i, 4) = 0 Then
                MsgBox "Ｙ軸設定（測定値）の分割数は、0より大きい値を入力してください。", , "ﾒｯｾｰｼﾞ"
                ckSW = False
                Exit For
            End If
            If CSng(F1Book1.TextRC(i, 2)) > CSng(F1Book1.TextRC(i, 3)) Then
                MsgBox "Ｙ軸設定の最小値の方が、最大値より大きい値が入力されています。確認してください。", , "ﾒｯｾｰｼﾞ"
                ckSW = False
                Exit For
            End If
            If CSng(F1Book1.TextRC(i, 2)) = CSng(F1Book1.TextRC(i, 3)) Then
                MsgBox "Ｙ軸設定の最小値と最大値に、同じ値が入力されています。確認してください。", , "ﾒｯｾｰｼﾞ"
                ckSW = False
                Exit For
            End If
        Next i
        If ckSW = False Then Exit Sub
        
        keiji.Xmax = Text1(0).Text
        keiji.xBUN = Text1(1).Text
        Call WriteIni("経時変化図設定", "表示時間数", CStr(keiji.Xmax), CurrentDIR & "計測設定.ini")
        Call WriteIni("経時変化図設定", "Ｘ軸分割数", CStr(keiji.xBUN), CurrentDIR & "計測設定.ini")
        
        SS(1) = "最小値"
        SS(2) = "最大値"
        SS(3) = "分割数"
        co = 0
        For i = 1 To SelKou(0)
            k_ID = SelKou(i)
            SS(0) = CStr(k_ID)
            For j = 1 To kou(k_ID, 0).no
                co = co + 1
                kou(k_ID, j).Kmin = F1Book1.TextRC(co, 2)
                kou(k_ID, j).Kmax = F1Book1.TextRC(co, 3)
                kou(k_ID, j).KBUN = F1Book1.TextRC(co, 4)
                If kou(k_ID, 0).no > 1 Then SS(0) = SS(0) & CStr(j)
                
                Call WriteIni("経時変化図設定", SS(1) & SS(0), CStr(kou(k_ID, j).Kmin), CurrentDIR & "計測設定.ini")
                Call WriteIni("経時変化図設定", SS(2) & SS(0), CStr(kou(k_ID, j).Kmax), CurrentDIR & "計測設定.ini")
                Call WriteIni("経時変化図設定", SS(3) & SS(0), CStr(kou(k_ID, j).KBUN), CurrentDIR & "計測設定.ini")
            Next j
        Next i
        
        For i = 1 To KeijiMAX
            Kkou(i).ck = Check1(i - 1).Value
            If Check1(i - 1).Value = 1 Then
                Kkou(i).kou = SetKou(i, 1, Combo1(i - 1).ListIndex + 1)
                Kkou(i).type = SetKou(i, 2, Combo1(i - 1).ListIndex + 1)
                
                Kkou(i).dan = Combo2(i - 1).ListIndex + 1
                If Kkou(i).dan = 0 And DanSet(Kkou(i).kou, 0).dan = 1 Then Kkou(i).dan = 1
                
                co = 0
                For j = 1 To List1(i - 1).ListCount
                    If List1(i - 1).Selected(j - 1) = True Then
                        co = co + 1
                        Kkou(i).ten(co) = j
                    End If
                Next j
                Kkou(i).ten(0) = co
            End If
        Next i
        f = FreeFile
        Open CurrentDIR & "keiji.dat" For Output Lock Write As #f
            For i = 1 To KeijiMAX
                SS(0) = CStr(i)
                SS(0) = SS(0) & ", " & CStr(Kkou(i).ck)
                If Kkou(i).ck = 1 Then
                    SS(0) = SS(0) & ", " & CStr(Kkou(i).dan)
                    SS(0) = SS(0) & ", " & CStr(Kkou(i).kou)
                    SS(0) = SS(0) & ", " & CStr(Kkou(i).type)
                    SS(0) = SS(0) & ", " & CStr(Kkou(i).ten(0))
                    For j = 1 To Kkou(i).ten(0)
                        SS(0) = SS(0) & ", " & CStr(Kkou(i).ten(j))
                    Next j
                End If
                Print #f, SS(0)
            Next i
        Close #f
        frmCLOSE.keijiSet = True
    End If
    Unload Me
End Sub

Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
Dim YMIN As Single, YMAX As Single
    
    If IsNumeric(EditString) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        CancelFlag = True
        'F1Book1.CancelEdit
        Exit Sub
    End If
    
    Select Case F1Book1.Col
    Case 4
        If CInt(EditString) = 0 Then
            MsgBox "Ｙ軸設定（測定値）の分割数は、0より大きい値を入力してください。", vbCritical, "ﾒｯｾｰｼﾞ"
            CancelFlag = True
            'F1Book1.CancelEdit
        End If
    Case Else
        If F1Book1.Col = 2 Then YMIN = CSng(EditString): YMAX = CSng(F1Book1.TextRC(F1Book1.Row, 3))
        If F1Book1.Col = 3 Then YMAX = CSng(EditString): YMIN = CSng(F1Book1.TextRC(F1Book1.Row, 2))
    
        If YMIN > YMAX Then
            MsgBox "Ｙ軸設定の最小値の方が、最大値より大きい値が入力されています。確認してください。", vbCritical, "ﾒｯｾｰｼﾞ"
            CancelFlag = True
            'F1Book1.CancelEdit
            Exit Sub
        End If
        If YMIN = YMAX Then
            MsgBox "Ｙ軸設定の最小値と最大値に、同じ値が入力されています。確認してください。", vbCritical, "ﾒｯｾｰｼﾞ"
            CancelFlag = True
            'F1Book1.CancelEdit
            Exit Sub
        End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer, SS As String
    Dim j As Integer, co As Integer
    Dim f As Integer, L As String
    Dim Hintv As Single, Mintv As Single, Sintv As Single
    Dim k_ID As Integer, t_ID As Integer
    
    Erase Kkou, SelKou, SetKou
    
    frmCLOSE.keijiSet = False
    stCK = False
    j = 0
    For k_ID = 1 To kou(0, 1).no - 1
        If k_ID <> 2 Then
            If DanSet(k_ID, 0).dan > 0 Then
                j = j + 1
                SelKou(j) = k_ID
            End If
        End If
    Next k_ID
    SelKou(0) = j
    
    Text1(0).Text = Format$(keiji.Xmax, "0")
    Text1(1).Text = Format$(keiji.xBUN, "0")

    If Second(KE_intv) = 0 Then Sintv = 1 Else Sintv = 60 / Second(KE_intv)
    If Minute(KE_intv) = 0 Then
        If Second(KE_intv) = 0 Then Mintv = 1 Else Mintv = 60
    Else
        Mintv = 60 / Minute(KE_intv)
    End If
    If Hour(KE_intv) = 0 Then Hintv = 24 Else Hintv = 24 / Hour(KE_intv)
    
    MaxH = 69120 \ (Hintv * Mintv * Sintv)
    Label2.Caption = "（最大時間＝" & CStr(MaxH) & "）"
    
    i = FileCheck(CurrentDIR & "keiji.dat", "経時変化図設定データ")
    If i = 0 Then Unload 計測Form: End
    
    f = FreeFile
    Open CurrentDIR & "keiji.dat" For Input As #f
        For i = 1 To KeijiMAX
            Input #f, L: Input #f, L
            Kkou(i).ck = CInt(L)
            If Kkou(i).ck = 1 Then
                Input #f, L: Kkou(i).dan = CInt(L)
                Input #f, L: Kkou(i).kou = CInt(L)
                Input #f, L: Kkou(i).type = CInt(L)
                Input #f, L: Kkou(i).ten(0) = CInt(L)
                For j = 1 To Kkou(i).ten(0)
                    Input #f, L
                    Kkou(i).ten(j) = CInt(L)
                Next j
            End If
        Next i
    Close #f
    
    For i = 1 To KeijiMAX
        Check1(i - 1).Value = Kkou(i).ck
        
        '項目
        Combo1(i - 1).Clear
        Combo2(i - 1).Clear
        For j = 1 To 20
            SetKou(i, 1, j) = 0
            SetKou(i, 2, j) = 0
        Next j
        
        co = 0
        For j = 1 To SelKou(0)
            k_ID = SelKou(j)
            
            For t_ID = 1 To kou(k_ID, 0).no
                co = co + 1
                SS = Trim$(kou(k_ID, t_ID).TI1)
                If kou(k_ID, 0).no > 1 Then SS = SS & " " & Trim$(kou(k_ID, t_ID).TI2)
                
                Combo1(i - 1).AddItem SS
                SetKou(i, 1, co) = k_ID
                SetKou(i, 2, co) = t_ID
                
                If Kkou(i).ck = 1 Then
                    If Kkou(i).kou = k_ID And Kkou(i).type = t_ID Then
                        Combo1(i - 1).ListIndex = co - 1
                    End If
                End If
            Next t_ID
        Next j
        If Kkou(i).ck = 0 Then Combo1(i - 1).ListIndex = 0
    Next i
    
    '測定値
    co = 0
    For i = 1 To SelKou(0)
        k_ID = SelKou(i)
        For j = 1 To kou(k_ID, 0).no
            co = co + 1
            SS = Trim$(kou(k_ID, j).TI1)
            If kou(k_ID, 0).no > 1 Then SS = SS & " " & Trim$(kou(k_ID, j).TI2)
            F1Book1.TextRC(co, 1) = SS
            F1Book1.TextRC(co, 2) = kou(k_ID, j).Kmin
            F1Book1.TextRC(co, 3) = kou(k_ID, j).Kmax
            F1Book1.TextRC(co, 4) = kou(k_ID, j).KBUN
        Next j
    Next i
    F1Book1.MaxRow = co
    F1Book1.SetActiveCell 1, 2
    F1Book1.EnableProtection = True
    F1Book1.DoSafeEvents = True
    F1Book1.Modified = False
    
    stCK = True
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim SW As Boolean

    SW = False
    
    If IsNumeric(Text1(Index).Text) = False Then
        MsgBox "数値と認識できない値が入力されました。もう一度、入力してください。", vbCritical, "エラーメッセージ"
        Text1(Index).SetFocus
        GoTo txt_skip
    End If
    If Text1(Index).Text <= 0 Then
        MsgBox "0より大きい数値を入力してください。", vbCritical, "エラーメッセージ"
        Text1(Index).SetFocus
        GoTo txt_skip
    End If
    
    If Index = 0 Then
        If Text1(0).Text > MaxH Then
            MsgBox "表示時間数は、最大時間より小さい値を入力してください。", vbCritical, "ﾒｯｾｰｼﾞ"
            Text1(Index).SetFocus
            GoTo txt_skip
        End If
    End If
    
    SW = True
    
txt_skip:
''    If SW = False Then
''        If Index = 0 Then Text1(0).Text = Format$(keiji.XMAX, "0")
''        If Index = 1 Then Text1(1).Text = Format$(keiji.Xbun, "0")
''    Else
''        If Index = 0 Then keiji.XMAX = Text1(0).Text
''        If Index = 1 Then keiji.Xbun = Text1(1).Text
''    End If
End Sub
