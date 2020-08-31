VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "Vsflex8l.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOset 
   Caption         =   "位置設定"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "frmPOset.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows の既定値
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '下揃え
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4185
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10372
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ハードコピー"
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   6840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7440
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command5 
         Caption         =   "読み込み"
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VSFlex8LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   5415
         _cx             =   5080
         _cy             =   5080
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         Caption         =   "読み込み位置"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   525
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   6120
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox Check2 
         Caption         =   "器械点座標再設定"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "初期値を置き換える"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   495
      Left            =   6240
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
      _cx             =   873
      _cy             =   873
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "印刷中..."
      AbortTextButton =   "ｷｬﾝｾﾙ"
      AbortTextDevice =   "出力先 %s(%s)"
      AbortTextPage   =   "%d ﾍﾟｰｼﾞ目を印刷中"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   -0.367647058823529
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   5
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   0
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "ﾍﾟｰｼﾞ全体(&P)|ﾍﾟｰｼﾞ幅(&W)|2ﾍﾟｰｼﾞ(&T)|ｻﾑﾈｲﾙ(&N)"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
End
Attribute VB_Name = "frmPOset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Readco As Long
Dim SaveSW As Boolean

Sub Posave()
    Dim f As Integer, i As Integer, j As Integer, L As String
    Dim SS1 As String, SS2 As String
    Dim syokidt(10) As Double
    Dim f1 As Integer, f2 As Integer
    
    f = FreeFile
    Open InitDT.PoFILE0 For Output Lock Write As #f
    For i = 1 To InitDT.Tensu
        If VSFlexGrid1.TextMatrix(i, 1) = "" Or VSFlexGrid1.TextMatrix(i, 1) = "******" Then
            PoDT.H(1, i) = -999
        Else
            PoDT.H(1, i) = VSFlexGrid1.TextMatrix(i, 1)
        End If
        If VSFlexGrid1.TextMatrix(i, 2) = "" Or VSFlexGrid1.TextMatrix(i, 2) = "******" Then
            PoDT.V(1, i) = -999
        Else
            PoDT.V(1, i) = VSFlexGrid1.TextMatrix(i, 2)
        End If
        If VSFlexGrid1.TextMatrix(i, 3) = "" Or VSFlexGrid1.TextMatrix(i, 3) = "******" Then
            PoDT.s(1, i) = -999
        Else
            PoDT.s(1, i) = VSFlexGrid1.TextMatrix(i, 3)
        End If
        
        SS1 = Format(i, "@@@@")
        SS2 = CStr(PoDT.H(1, i))
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        SS2 = CStr(PoDT.V(1, i))
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        SS2 = CStr(PoDT.s(1, i))
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        Print #f, SS1
    Next i
    Close #f

    FileCopy InitDT.PoFILE0, InitDT.PoFILE1

    
    For i = 1 To Tbl(1, 1, 0).ten
        syokidt(i) = -999999
        j = (Tbl(1, 1, i).FLD + 2) \ 3
        If PoDT.s(1, j) <> -999 Then syokidt(i) = ZD(j)
    Next i
    
    If Check1.Value = 1 Then
        i = 0
        f2 = FreeFile
        Open CurrentDir & "ctable_dm.dat" For Output As #f2
        f1 = FreeFile
        Open Tabl_path & "ctable.dat" For Input Shared As #f1
        Do While Not (EOF(f1))
            Line Input #f1, L
            If Left$(L, 1) = ";" Then
                Print #f2, L
            Else
                i = i + 1
                If syokidt(i) = -999999 Then
                    Print #f2, L
                Else
                    SS2 = Format(syokidt(i), "0.000")
                    SS1 = Left(L, 16) & Space$(10 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
                    SS1 = SS1 & Right(L, 18)
                    Print #f2, SS1
                End If
            End If
        Loop
        Close #f1
        Close #f2
        If Dir(Tabl_path & "ctable.dat") <> "" Then Kill Tabl_path & "ctable.dat"
        FileCopy CurrentDir & "ctable_DM.dat", Tabl_path & "ctable.dat"
    End If
    
    f = FreeFile
    Open CurrentDir & "Zhyou-Syoki.dat" For Output Lock Write As #f
    Print #f, Format(Now, "yyyy/mm/dd hh:nn:ss")
    For i = 1 To InitDT.Tensu
        Print #f, XD(i) & " , ";
        Print #f, YD(i) & " , ";
        Print #f, ZD(i)
    Next i
    Close #f

    If Check2.Value = 1 Then
        Call frmGTS800A.CalcAzimuth
        InitDT.x0 = -XD(1) * Cos(InitDT.AZIMUTH * RAD#)
        InitDT.y0 = -XD(1) * Sin(InitDT.AZIMUTH * RAD#)
        InitDT.z0 = -ZD(1)
        InitDT.MH = 0
        InitDT.x1 = 0
        InitDT.y1 = 0
        InitDT.z1 = 0
        
        With frmSyokiset
            .xTextN1(1).Text = InitDT.x0
            .xTextN1(2).Text = InitDT.y0
            .xTextN1(3).Text = InitDT.z0
            .xTextN1(4).Text = InitDT.MH
            .xTextN1(5).Text = InitDT.x1
            .xTextN1(6).Text = InitDT.y1
            .xTextN1(7).Text = InitDT.z1
            .xTextN1(8).Text = InitDT.AZIMUTH
        End With
    End If

    SaveSW = False
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim Response As Integer
    
    If Index = 0 Then
        If SaveSW = True Then
            Response = MsgBox("変更が保存されていません。保存しますか？", vbYesNoCancel + vbExclamation, "終了の確認")
            If Response = vbCancel Then Exit Sub
            If Response = vbYes Then
                Call Posave
            End If
        End If
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    Call Posave
End Sub

Private Sub Command3_Click()
    Dim Winh As Long
    Dim DC As Long
    Dim Wrect As RECT
    Dim Ret As Long
    Dim ppp As Object
    Set ppp = Picture
    
    Winh = GetForegroundWindow  'Winh = GetActiveWindow
    DC = GetWindowDC(Winh)
    Ret = GetWindowRect(Winh, Wrect)
    
    Picture1.Width = Me.Width
    Picture1.Height = Me.Height
    
    Ret = BitBlt(Picture1.hdc, 0, 0, Wrect.Right, Wrect.Bottom, DC, 0, 0, SRCCOPY)
    Ret = ReleaseDC(Winh, DC)
    
    'Clipboard.Clear
    'Clipboard.SetData Picture1.Image
    
    
    With VSPrinter1
        .Zoom = 100
        .Orientation = orLandscape
        .StartDoc
        .DrawPicture Picture1.Image, "1in", "1in", "90%", "90%"
        .EndDoc
        .Action = paChoosePrintAll
        '.PrintDoc
    End With

End Sub

Private Sub Command5_Click()
    Dim no As Integer
    Dim h1#, v1#, s1#, c%, rl$
    Dim i%
    
    no = Combo1.ListIndex + 1
    
    If no = 1 Then
        MsgBox "原点方向を正で視準し、ENTERを押してください。", vbOKOnly, "ﾒｯｾｰｼﾞ"
    ElseIf no = 2 Then
        MsgBox "基準点方向を正で視準し、ENTERを押してください。", vbOKOnly, "ﾒｯｾｰｼﾞ"
    Else
        MsgBox CStr(no - 2) & "方向を正で視準し、ENTERを押してください。", vbOKOnly, "ﾒｯｾｰｼﾞ"
    End If
    
    Readco = Readco + 1

GoTo 12
    If no = 1 Then Call frmGTS800A.SendCmd("ZB1" + "+0000000d")     '水平角の設定
    
    Call frmGTS800A.DataIn(h1#, v1#, s1#, c%, rl$)
    VSFlexGrid1.TextMatrix(no, 1) = h1#
    VSFlexGrid1.TextMatrix(no, 2) = v1#
    VSFlexGrid1.TextMatrix(no, 3) = s1#
GoTo 14
12
    h1 = VSFlexGrid1.TextMatrix(no, 1) '= h1#
    v1 = VSFlexGrid1.TextMatrix(no, 2) '= v1#
    s1 = VSFlexGrid1.TextMatrix(no, 3) '= s1#
    
14
    If (s1# <> 0) Then
        Call frmGTS800A.XyzCal(no, h1#, v1#, s1#, rl$)
    End If
    
    SaveSW = True
    
    Debug.Print XD(no), YD(no), ZD(no)

    MsgBox "計測が完了しました。", vbInformation, "ﾒｯｾｰｼﾞ"
End Sub

Private Sub Form_Load()
    Dim f As Integer, L As String, i As Integer
    Dim ss As String
    Dim rc As Integer
    
    Top = frmSyokiset.Top + frmSyokiset.Height - 1000
    Left = frmSyokiset.Left + frmSyokiset.Width - 1000
    
    Timer1.Interval = 500
    
'    For i = 1 To InitDT.Tensu
'        Combo1.AddItem "No." & CStr(i)
'    Next i
'    Combo1.ListIndex = 0
    
    VSFlexGrid1.Rows = InitDT.Tensu + 1
    VSFlexGrid1.ColWidth(0) = 1000
    VSFlexGrid1.ColWidth(1) = 1400
    VSFlexGrid1.ColWidth(2) = 1400
    VSFlexGrid1.ColWidth(3) = 1400
    VSFlexGrid1.ColAlignment(0) = 4
    For i = 1 To VSFlexGrid1.Cols - 1
        VSFlexGrid1.ColAlignment(i) = 7
        VSFlexGrid1.Row = 0: VSFlexGrid1.Col = i: VSFlexGrid1.CellAlignment = 4
    Next i
    VSFlexGrid1.TextMatrix(0, 1) = "H (秒)"
    VSFlexGrid1.TextMatrix(0, 2) = "V (秒)"
    VSFlexGrid1.TextMatrix(0, 3) = "S (ｍ)"
    
    For i = 1 To 10
        PoDT.H(1, i) = -999
        PoDT.V(1, i) = -999
        PoDT.s(1, i) = -999
    Next i
    
    For i = 1 To InitDT.Tensu
        If i = 1 Then
            ss = "原点方向"
        ElseIf i = 2 Then
            ss = "基準方向"
        Else
            ss = "No." & CStr(i - 2)
        End If
        VSFlexGrid1.TextMatrix(i, 0) = ss
        Combo1.AddItem ss
    Next i
    Combo1.ListIndex = 0
    
    If Dir(InitDT.PoFILE0) <> "" Then
        i = 0
        f = FreeFile
        Open InitDT.PoFILE0 For Input Shared As #f
        Do While Not (EOF(f))
            Line Input #f, L
            i = i + 1
            If i > InitDT.Tensu Then Exit Do
            PoDT.H(1, i) = CDbl(Mid(L, 5, 12))
            PoDT.V(1, i) = CDbl(Mid(L, 17, 12))
            PoDT.s(1, i) = CDbl(Mid(L, 29, 12))
            
            If PoDT.H(1, i) = -999 Then
                VSFlexGrid1.TextMatrix(i, 1) = "******"
            Else
                VSFlexGrid1.TextMatrix(i, 1) = PoDT.H(1, i)
            End If
            If PoDT.V(1, i) = -999 Then
                VSFlexGrid1.TextMatrix(i, 2) = "******"
            Else
                VSFlexGrid1.TextMatrix(i, 2) = PoDT.V(1, i)
            End If
            If PoDT.s(1, i) = -999 Then
                VSFlexGrid1.TextMatrix(i, 3) = "******"
            Else
                VSFlexGrid1.TextMatrix(i, 3) = PoDT.s(1, i)
            End If
        Loop
        Close #f
    End If
    
    Check1.Value = 1
    Check2.Value = 1
    Readco = 0
    SaveSW = False
    
        InitDT.x0 = 0 '-XD(1) * Cos(InitDT.AZIMUTH * RAD#)
        InitDT.y0 = 0 '-XD(1) * Sin(InitDT.AZIMUTH * RAD#)
        InitDT.z0 = 0 '-ZD(1)
        InitDT.AZIMUTH = 0
        
'    With frmGTS800A
'        .VBMCom1.VcDeviceName = RsInit.DeviceNo
'        .VBMCom1.VcBaudRate = RsInit.SpdNO
'        .VBMCom1.VcParity = RsInit.PrtNO
'        .VBMCom1.VcByteSize = RsInit.sizeNO
'        .VBMCom1.VcStopBits = RsInit.stopNo
'        .VBMCom1.VcRecvTimeOut = RsInit.Rtime
'        .VBMCom1.VcSendTimeOut = RsInit.Stime
'    End With
'
'    '通信ポートをオープン
'    rc = frmGTS800A.VBMCom1.OpenComm
'    If rc <> 0 Then
'      MsgBox "通信ポートがオープンできません。" & CStr(rc), vbCritical
'      End
'    End If
    
'    Call GTS8init
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rc As Integer
    
    '通信ポートをクローズ
    rc = frmGTS800A.VBMCom1.CloseComm

End Sub

Private Sub Timer1_Timer()
    Me.StatusBar1.Panels(2).Text = Format(Now, "yyyy/mm/dd hh:nn:ss")
End Sub


