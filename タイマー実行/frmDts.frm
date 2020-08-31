VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDts 
   Caption         =   "データシート"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   Icon            =   "frmDts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11175
   Begin VB.CommandButton Command1 
      Caption         =   "印刷"
      Height          =   495
      Index           =   1
      Left            =   9360
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出力日変更"
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "ページ"
      Height          =   1695
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
      Begin VB.CommandButton cmdPage 
         Caption         =   "≪先頭"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "First Page"
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton cmdPage 
         Caption         =   "最終≫"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Last Page"
         Top             =   1200
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.HScrollBar scrlPage 
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Max             =   1
         Min             =   1
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Value           =   1
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  '不透明
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ズーム"
      Height          =   1815
      Index           =   0
      Left            =   8880
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton cmdZoom 
         Caption         =   "縮小"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdZoom 
         Caption         =   "拡大"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '中央揃え
         BackStyle       =   0  '透明
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  '塗りつぶし
         Height          =   285
         Index           =   0
         Left            =   1080
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "表示倍率"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "メニューに戻る"
      Height          =   495
      Index           =   0
      Left            =   9360
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4995
      _cx             =   8819
      _cy             =   8819
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
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
      Zoom            =   28.125
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
   Begin VB.Menu mnuFile 
      Caption         =   "ﾌｧｲﾙ"
      Visible         =   0   'False
      Begin VB.Menu mnuPrinterSet 
         Caption         =   "ﾌﾟﾘﾝﾀｰ設定"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "印刷"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "ﾌｧｲﾙへ保存"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "終了"
      End
   End
End
Attribute VB_Name = "frmDts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------
'   データシートを表示・印刷・ファイル保存する
'
'       印刷設定 サイズ＝Ａ４
'                向き＝横
'                文字フォント＝ＭＳ 明朝
'                余白＝上25mm、下20mm、左10mm、右10mm
'----------------------------------------------------------------------------------------------
'出力先 1.表示  2.ファイル保存
Dim PrnMode As Integer

'表示作業確認 True=作業中 False=作業完了
Public HYOUsw As Boolean

'True=作業中に中断された場合
Dim Tyuudan As Boolean

'項目番号・断面番号
Dim kou_ID As Integer, dan_ID As Integer

'種類番号
Dim HENI As Integer

'出力先ｵﾌﾞｼﾞｪｸﾄ
Private TARGETOBJECT As Object

'データシート行文字列
Dim strBody  As String

'カレントパス
Dim CuDir As String

'パラメータ
Dim DS_SD As Date, DS_ED As Date '開始日・終了日
Dim SDtype As Integer            'データ条件
Dim STtype As Integer            '先頭時間
Dim SEEKtime As Integer          '作表条件用 抽出時間
Dim SEEKMday(24) As String       '    〃    １日でどの時間を作図するか時間を代入する
Dim sTYPE As Integer             '表示形式 0.物理量 1.測定値
Dim DTS_Col_WIDTH As Integer     '列の幅
Dim DTS_Col_MAX   As Integer     '列数

'**********************************************************************************************
'   作表処理（画面描画）
'**********************************************************************************************
Public Sub Sakuhyou()
    Dim Mdate As Date
    Dim i As Integer, f As Integer, j As Integer, jj As Integer
    Dim pd(50) As Double
    Dim bf As String
    Dim SW As Boolean
    Dim po As Long
    Dim fmt As String
    Dim FLDno As Integer, FLDco As Integer, FLDstep As Integer
    Dim FileName As String
    Dim pdSNG As Single, pdDBL As Double
    Dim Kpd As Double
    
    DS_SD = Hsheet.SD
    DS_ED = Hsheet.ED
    
    Screen.MousePointer = 11
    Me.WindowState = 0
    
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    mnuPrint.Enabled = False
    mnuFileSave.Enabled = False
    
    HYOUsw = True
    
    '印刷開始
    If PrnMode = 1 Then VSPrinter1.Visible = False
    If PrnMode = 1 Then VSPrinter1.StartDoc
    
    FileName = KEISOKU.Data_path & DATA_DAT

    If Dir(FileName) = "" Then
        Open FileName For Output As #1
        Close #1
    End If

    po = STARTpoint(DS_SD)

    f = FreeFile
    Open FileName For Input Access Read Shared As #f
        Seek #f, po
        Do While Not (EOF(f))

            For i = 1 To Tbl(kou_ID, dan_ID, 0).ten: pd(i) = 999999: Next i

            Line Input #f, bf

            Mdate = CDate(Mid$(bf, 1, 19))

            If Mdate > DS_ED Then Exit Do
            If Mdate < DS_SD Then GoTo noDts

            If 0 < SDtype Then
                For i = 1 To 24 / SEEKtime
                    If Format(Mdate, "hh:nn:ss") = SEEKMday(i) Then Exit For
                Next i
                If i > (24 / SEEKtime) Then GoTo noDts
            End If

            DoEvents

            For i = 1 To Tbl(kou_ID, dan_ID, 0).ten
                FLDno = 20 + 10 * (Tbl(kou_ID, dan_ID, i).FLD - 1): FLDco = 10
                If IsNumeric(Mid$(bf, FLDno, FLDco)) = True Then
                    pdDBL = CDbl(Mid$(bf, FLDno, FLDco))
                Else
                    pdDBL = 999999
                End If

                If Abs(pdDBL) >= 999999 Or Tbl(kou_ID, dan_ID, i).Syo = 999999 Then
                    pd(i) = 999999
                Else
                    '2001/12/11
                    pd(i) = (pdDBL - Tbl(kou_ID, dan_ID, i).Syo) * Tbl(kou_ID, dan_ID, i).Kei
''                    pd(i) = (-1) * (pdDBL - Tbl(kou_ID, dan_ID, i).Syo) * Tbl(kou_ID, dan_ID, i).Kei
                End If
            Next i
            Call KEISAN(pd(), Kpd)


            If kou(kou_ID, HENI).dec = 0 Then
                fmt = "#0"
            Else
                fmt = "#0." & String$(kou(kou_ID, HENI).dec, "0")
            End If

            strBody = Mid$(bf, 1, 19)
            
            If PrnMode = 2 Then strBody = strBody & ","
            If Abs(Kpd) >= 999999 Then
                strBody = strBody & Format$(String(DTS_Col_WIDTH - 2, "*"), String(DTS_Col_WIDTH, "@"))
            Else
                strBody = strBody & Format$(Format$(Kpd, fmt), String(DTS_Col_WIDTH, "@"))
            End If
            
            For i = 1 To Tbl(kou_ID, dan_ID, 0).ten
                If PrnMode = 2 Then strBody = strBody & ","

                pd(0) = pd(i)

                If Abs(pd(0)) >= 999999 Then
                    strBody = strBody & Format$(String(DTS_Col_WIDTH - 2, "*"), String(DTS_Col_WIDTH, "@"))
                Else
                    strBody = strBody & Format$(Format$(pd(0), fmt), String(DTS_Col_WIDTH, "@"))
                End If
            Next i

            If PrnMode = 1 Then
                With VSPrinter1
                    .Paragraph = strBody
                End With
            Else
                Print #3, strBody
            End If
noDts:
        Loop
    Close #f
    
    If PrnMode = 1 Then VSPrinter1.EndDoc             ' 印刷を終了します。
    
    mnuPrint.Enabled = True
    mnuFileSave.Enabled = True
    If PrnMode = 1 Then
        Command1(1).Enabled = True
        Command1(2).Enabled = True
        VSPrinter1.Visible = True
        cmdPage(0).Visible = True: cmdPage(1).Visible = True
        cmdZoom(0).Visible = True: cmdZoom(1).Visible = True
        Label1(0).Visible = True: Label1(1).Visible = True
        scrlPage.Visible = True
    End If
    
    Screen.MousePointer = 0
    HYOUsw = False

End Sub

'**********************************************************************************************
'   表示ページ移動
'       Index 0=先頭 1=最終
'**********************************************************************************************
Private Sub cmdPage_Click(Index As Integer)

    If Index = 0 Then
        VSPrinter1.PreviewPage = 1
        scrlPage.Value = 1
    Else
        VSPrinter1.PreviewPage = VSPrinter1.PageCount
        scrlPage.Value = VSPrinter1.PageCount
    End If
    
    Label1(1).Caption = Format$(VSPrinter1.PreviewPage) & "/" & Format$(VSPrinter1.PageCount) & " ﾍﾟｰｼﾞ"

End Sub

'**********************************************************************************************
'   表示されてるデータシートの縮尺を設定します。
'       縮尺ｲﾝﾀｰﾊﾞﾙ：ZoomParam
'       最大縮尺率：150
'       最小縮尺率：20
'**********************************************************************************************
Private Sub cmdZoom_Click(Index As Integer)
    Const ZoomParam = 10
    
    With VSPrinter1
        Select Case Index
            Case 0
                .Zoom = .Zoom + ZoomParam
                'Zoomの許容範囲を外れる場合、[拡大]ボタンを使用不可能に設定
                If .Zoom > 150 - ZoomParam Then
                    cmdZoom(0).Enabled = False
                End If
                '[縮小]ボタンを使用可能に設定
                If Not cmdZoom(1).Enabled Then
                    cmdZoom(1).Enabled = True
                End If
            Case 1
                .Zoom = .Zoom - ZoomParam
                'Zoomの許容範囲を外れる場合、[縮小]ボタンを使用不可能に設定
                If .Zoom <= 0 + ZoomParam Then
                    cmdZoom(1).Enabled = False
                End If
                '[拡大]ボタンを使用可能に設定
                If Not cmdZoom(0).Enabled Then
                    cmdZoom(0).Enabled = True
                End If
        End Select
        'Me.Caption = "Zoom " & .Zoom & "%"
    End With

    Label1(0).Caption = Format$(VSPrinter1.Zoom) & "%"

End Sub


Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then Unload Me
    If Index = 1 Then VSPrinter1.Action = paChoosePrintAll
    If Index = 2 Then frmDTSpara.Show
End Sub

'**********************************************************************************************
'   フォームの初期設定
'   VSPrinter1の初期設定
'**********************************************************************************************
Private Sub Form_Load()
    Dim i As Integer, j As Integer
    Dim STtime As Date
    Dim ss As String
    Dim f As Integer
    Dim L As String
    '14940         19185
    Me.Height = Screen.Height - 420 '11000 '16590
    Me.Width = Screen.Width '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
    
    CuDir = CurDir
    
'    Set TARGETOBJECT = VSPrinter1
    mnuPrint.Enabled = True
    
'    Me.Caption = "データシート"
'    Me.Height = 10840
'    Me.Width = 13785
'    Me.top = 0
'    Me.Left = Screen.Width - Width
    
    HYOUsw = False
    Tyuudan = False
    PrnMode = 1
            
    PrntDrvSW = True
    
    With VSPrinter1
        .Visible = False
        
        '印刷可能領域を表示
        .ShowGuides = gdShow: 'Me.mnuSGuides.Checked = True
        'マウスをドラッグすることによりページのプレビューをスクロール
'        .MouseScroll = False
'        .MouseZoom = False
        
        '各ページの周りに描かれるページ枠を設定
        .PageBorder = 0  'pbAll
        'Printerコントロールの出力を全て画面へ
        .Preview = True
        'プレビュー画面の縮尺率
        .Zoom = 100
        .ZoomMode = zmPercentage
        '用紙サイズをＡ４に設定
        If .PaperSizes(vbPRPSA4) = True Then
            .PaperSize = pprA4
        Else
            PrntDrvSW = False
            MsgBox "用紙サイズを設定できませんでした。", vbExclamation
            Screen.MousePointer = vbDefault   'マウスを既定値に戻す
            Exit Sub
        End If
        '用紙方向を横に設定
        .Orientation = orPortrait
        If .Error <> 0 Or .Orientation = orLandscape Then
            PrntDrvSW = False
            MsgBox "用紙方向を設定できませんでした。", vbExclamation
            Screen.MousePointer = vbDefault   'マウスを既定値に戻す
            Exit Sub
        End If
    End With
    
    With VSPrinter1
        .MarginTop = "20mm"    '0
        .MarginBottom = "20mm" '720
        .MarginLeft = "25mm"   '1080 '1440
        .MarginRight = "10mm"  '540
    End With
    
    kou_ID = 1: dan_ID = 1: HENI = 1
    DTS_Col_MAX = Tbl(kou_ID, dan_ID, 0).ten + 1
    DTS_Col_WIDTH = 10
    
    SDtype = 0
    STtype = 0
    
    f = FreeFile
    Open KEISOKU.Data_path & DATA_DAT For Input Shared As #f
        If LOF(f) > 0 Then
            Line Input #f, L: DS_SD = CDate(Mid$(L, 1, 19))
        End If
        If LOF(f) - REC_LEN * 2 > 0 Then
            Seek #f, LOF(f) - REC_LEN * 2
        End If
        Do While Not (EOF(f))
            Line Input #f, L
        Loop
    Close #1
    If L <> "" Then DS_ED = CDate(Mid$(L, 1, 19))
    
    Hsheet.SD = DS_SD
    Hsheet.ED = DS_ED
End Sub

'**********************************************************************************************
'   作業中に中断した場合は、印刷ジョブを削除します。
'**********************************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If HYOUsw = True Then
        Tyuudan = True
        If PrnMode = 1 Then VSPrinter1.KillDoc
    End If
    frmMenu.Show
End Sub

'**********************************************************************************************
'   フォームのサイズを変更した場合にコントロールの位置を設定
'**********************************************************************************************
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
'    With cmdZoom(0)
'        .Left = mintControlMargin
'        .top = 75 'Me.ScaleHeight - mintControlMargin - .Height
'    End With
'    With cmdZoom(1)
'        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin / 2
'        .top = 75 'Me.ScaleHeight - mintControlMargin - .Height
'    End With
'    With Label1(1)
'        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin + cmdZoom(1).Width + mintControlMargin
'        .top = 175 'Me.ScaleHeight - mintControlMargin - cmdZoom(1).Height + 100
'    End With
'
'    With cmdPage(0)
'        .Left = Me.ScaleWidth - mintControlMargin / 2 * 4 - .Width - cmdPage(1).Width - Label1(0).Width - scrlPage.Width - 2000
'        .top = 75 'Me.ScaleHeight - mintControlMargin - .Height
'    End With
'    With cmdPage(1)
'        .Left = Me.ScaleWidth - mintControlMargin / 2 * 4 - .Width - Label1(0).Width - 2000
'        .top = 75 'Me.ScaleHeight - mintControlMargin - .Height
'    End With
'    With scrlPage
'        .Left = Me.ScaleWidth - mintControlMargin / 2 * 4 - .Width - cmdPage(1).Width - Label1(0).Width - 2000
'        .top = 75 'Me.ScaleHeight - mintControlMargin - .Height
'    End With
'    With Label1(0)
'        .Left = Me.ScaleWidth - mintControlMargin / 2 - .Width - 2000
'        .top = 175 'Me.ScaleHeight - mintControlMargin - cmdZoom(1).Height + 100
'    End With
'
'    With Command1
'        .Left = Me.ScaleWidth - 1680
'        .top = 75
'    End With
'    With VSPrinter1
'        .Left = mintControlMargin
'        .top = 480 'mintControlMargin
'        .Height = Me.ScaleHeight - cmdZoom(0).Height - 3 * mintControlMargin
'        .Width = Me.ScaleWidth - 2 * mintControlMargin
'    End With
    With Command1(0)
        .Left = Me.ScaleWidth - 1500
        .Top = 120
    End With
    With Command1(1)
        .Left = Me.ScaleWidth - 1500
        .Top = 960
    End With
    With Command1(2)
        .Left = Me.ScaleWidth - 1500
        .Top = 1800
    End With
    With Frame1(1)
        .Left = Me.ScaleWidth - 2060
        .Top = 2520 'Me.ScaleHeight - 2135
    End With
    With Frame1(0)
        .Left = Me.ScaleWidth - 2060
        .Top = 4440 'Me.ScaleHeight - 2135
    End With

    With VSPrinter1
        .Left = mintControlMargin
        .Top = mintControlMargin
        .Height = Me.ScaleHeight - 2 * mintControlMargin
        .Width = Me.ScaleWidth - 2 * mintControlMargin - 2100
    End With
Debug.Print Me.ScaleHeight, Me.ScaleWidth, Me.Height, Me.Width
End Sub

Private Sub HScroll1_Change()

End Sub

'**********************************************************************************************
'   メニュー〔ファイル〕〔終了〕
'**********************************************************************************************
Private Sub mnuEnd_Click()
    Unload Me
End Sub

'**********************************************************************************************
'   メニュー〔ファイル〕〔ファイルへ保存〕
'**********************************************************************************************
Private Sub mnuFileSave_Click()
'    Dim Datafile As String
'    Dim i As Integer, j As Integer
'    Dim SS As String, SS2 As String
'
'    CommonDialog1.FileName = "dts.csv"
'
'    CommonDialog1.CancelError = True ' CancelError プロパティを真 (True) に設定します。
'    On Error GoTo ErrHandler
'
'    CommonDialog1.Flags = cdlOFNFileMustExist        ' Flags プロパティを設定します。
'    CommonDialog1.Filter = "ＣＳＶ(カンマ区切り)|*.csv|" ' リスト ボックスに表示されるフィルタを設定します。
'    CommonDialog1.FilterIndex = 0                    ' "テキスト ファイル" を既定のフィルタとして指定します。
'    CommonDialog1.ShowSave                           ' [ファイルを開く] ダイアログ ボックスを表示します。
'    On Error GoTo 0
'
'    ' ユーザーが選択したファイル名を表示します。
'    Datafile = CommonDialog1.FileName
'
'    PrnMode = 2
'    ChDrive CuDir
'    ChDir CuDir
'    Open Datafile For Output As #3
'
'        SS = Trim$(kou(kou_ID, 1).TI1)
'        If kou(kou_ID, 0).no > 1 Then SS = SS & " " & Trim$(kou(kou_ID, HENI).TI2)
'        SS = SS & " データシート"
'        Print #3, SS
'        Print #3, TNAME1 & " " & TNAME2
'
'        If Trim$(DanSet(kou_ID, 0).dan) <> "" Then
'            Print #3, "断 面 ： " & Trim$(DanSet(kou_ID, dan_ID).ti)
'        Else
'            Print #3, ""
'        End If
'
'        Print #3, "単 位 ： " & Trim$(kou(kou_ID, HENI).Yt) & " (" & Trim$(kou(kou_ID, HENI).Yu) & ")"
'        Print #3, ""
'
'        '項目
'        SS = "    計 測 日 時    "
'        For i = 1 To Tbl(kou_ID, dan_ID, 0).ten
'            If Tbl(kou_ID, dan_ID, i).Sheet = 1 Then
'                SS2 = Trim$(Tbl(kou_ID, dan_ID, i).HAN)
'                SS = SS & "," & SS2
'            End If
'        Next i
'        Print #3, SS
'
'        Call Sakuhyou
'    Close #3
'
'    MsgBox "保存が完了しました。", vbInformation
'    Exit Sub
'ErrHandler:
'    ChDrive CuDir
'    ChDir CuDir
End Sub

'**********************************************************************************************
'   メニュー〔ファイル〕〔印刷〕
'**********************************************************************************************
Private Sub mnuPrint_Click()
    
    PrnMode = 1
    
    '印刷
    VSPrinter1.Action = paChoosePrintAll
End Sub

'**********************************************************************************************
'   メニュー〔ファイル〕〔ﾌﾟﾘﾝﾀｰ設定〕
'**********************************************************************************************
Private Sub mnuPrinterSet_Click()
    VSPrinter1.Action = paChoosePrinter
End Sub

'**********************************************************************************************
'   スクロール バーを使って、ページをスクロールさせ、その範囲内での現在位置を示すことができます。
'**********************************************************************************************
Private Sub scrlPage_Change()
    Dim lp As Integer
    
    scrlPage.SmallChange = VSPrinter1.PreviewPages
    scrlPage.LargeChange = scrlPage.SmallChange
    
    VSPrinter1.PreviewPage = scrlPage.Value
    
    lp = VSPrinter1.PreviewPage + VSPrinter1.PreviewPages - 1
    If lp > VSPrinter1.PageCount Then lp = VSPrinter1.PageCount
    If lp < VSPrinter1.PreviewPage Then lp = VSPrinter1.PreviewPage
    
    Label1(1).Caption = Format$(VSPrinter1.PreviewPage) & "/" & Format$(VSPrinter1.PageCount) & " ﾍﾟｰｼﾞ"

End Sub

'**********************************************************************************************
'   作表作業終了後、コントロールを再設定
'**********************************************************************************************
Private Sub VSPrinter1_EndDoc()
    
    VSPrinter1.PreviewPage = 1
    
    cmdZoom(0).Enabled = True
    cmdZoom(1).Enabled = True
    cmdPage(0).Enabled = True
    cmdPage(1).Enabled = True
    scrlPage.Enabled = True
    scrlPage.max = VSPrinter1.PageCount
    scrlPage.Value = VSPrinter1.PreviewPage
    scrlPage_Change
    Label1(0).Caption = Format$(VSPrinter1.Zoom) & "%"

End Sub

'**********************************************************************************************
'   ヘッダー表示
'**********************************************************************************************
Private Sub VSPrinter1_NewPage()
    Dim i As Integer, j As Integer, mco As Integer
    Dim ss As String, SS2 As String
    Dim haba1 As Single, haba2 As Single, maxcol As Integer

    '表の出力位置を設定

    With VSPrinter1
        .CurrentY = 1440
        .FontBold = False
        .FontItalic = False
        .FontSize = 9
        
        maxcol = DTS_Col_MAX
        haba1 = .TextWidth(String(19 + maxcol * DTS_Col_WIDTH, "-"))
        
        .FontBold = True
        .FontItalic = False
        .FontName = "ＭＳ 明朝"  'Courier
        .TextColor = vbBlack 'vbWhite
        .FontSize = 13
        
        ss = Trim$(kou(kou_ID, 1).TI1)
        ss = ss & "計測データシート"

        haba2 = .TextWidth(ss)

        If (haba1 - haba2) > 0 Then
            .IndentLeft = (haba1 - haba2) / 2
        Else
            .IndentLeft = 0
        End If
        .TextAlign = taLeftTop
        .Paragraph = ss

        .IndentLeft = 0
        .IndentRight = 0
        .Paragraph = ""

'        .FontSize = 11
'        .TextAlign = taLeftTop
'        .Paragraph = TNAME1 & " " & TNAME2

        .FontSize = 9
        .TextAlign = taLeftTop
'        .Paragraph = ""

        '.Paragraph = "単 位 ： " & Trim$(kou(kou_ID, HENI).Yt) & " (" & Trim$(kou(kou_ID, HENI).Yu) & ")"
        .Paragraph = "単 位 ： " & Trim$(kou(kou_ID, HENI).Yu)
        .Paragraph = ""

        .FontBold = False  ''.FontBold = True
        .FontItalic = False
        '.FontName = "ＭＳ ゴシック"  'Courier
        .FontSize = 9 '10
        .TextColor = vbBlack 'vbWhite
        '項目

        ss = "    計 測 日 時    "
        SS2 = "換算変位"
        mco = LenB(StrConv(SS2, vbFromUnicode))
        ss = ss & Space$(DTS_Col_WIDTH - mco) & SS2
        
        For i = 1 To Tbl(kou_ID, dan_ID, 0).ten
            SS2 = Trim$(Tbl(kou_ID, dan_ID, i).HAN)
            mco = LenB(StrConv(SS2, vbFromUnicode))
            ss = ss & Space$(DTS_Col_WIDTH - mco) & SS2
        Next i
        .Paragraph = ss
        
        .TextColor = vbBlack
        .FontBold = False
        .FontItalic = False
        .FontSize = 9
    End With

End Sub


