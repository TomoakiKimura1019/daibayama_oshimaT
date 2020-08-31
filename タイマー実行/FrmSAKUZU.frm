VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSakuzu 
   AutoRedraw      =   -1  'True
   Caption         =   "作図"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   1095
   ClientWidth     =   14880
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSAKUZU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   14880
   Begin VB.CommandButton Command1 
      Caption         =   "スケール変更"
      Height          =   495
      Index           =   2
      Left            =   7800
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "印刷"
      Height          =   495
      Index           =   1
      Left            =   7800
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "ズーム"
      Height          =   2055
      Left            =   7800
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
      Begin VB.CommandButton cmdZoom 
         Caption         =   "拡大"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdZoom 
         Caption         =   "縮小"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Visible         =   0   'False
         Width           =   1035
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
         Left            =   600
         TabIndex        =   8
         Top             =   510
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  '塗りつぶし
         Height          =   285
         Left            =   480
         Top             =   480
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
         TabIndex        =   7
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "メニューに戻る"
      Height          =   495
      Index           =   0
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1380
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7695
      _cx             =   13573
      _cy             =   11880
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
      Zoom            =   39.1098484848485
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
   Begin VB.Menu FAIRU 
      Caption         =   "ファイル"
      Visible         =   0   'False
      Begin VB.Menu mnuPrinterSet 
         Caption         =   "ﾌﾟﾘﾝﾀｰ設定"
      End
      Begin VB.Menu INSATU 
         Caption         =   "印刷"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu BITMAP 
         Caption         =   "ﾌｧｲﾙへ保存"
      End
      Begin VB.Menu mnueBar2 
         Caption         =   "-"
      End
      Begin VB.Menu SYUURYOU 
         Caption         =   "終了"
      End
   End
End
Attribute VB_Name = "frmSakuzu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------------------------------------------------------------------
'   経時変化図・断面分布図を表示・印刷・ファイル保存する
'
'       印刷設定 サイズ＝Ａ４
'                向き＝横
'                文字フォント＝ＭＳ ゴシック
'----------------------------------------------------------------------------------------------

'1.経時変化図 2.断面分布図
Public KDSW As Integer

'表示作業確認 True=作業中 False=作業完了
Public ZuOut As Boolean

'True=作業中に中断された場合
Dim Tyuudan As Boolean

'出力先ｵﾌﾞｼﾞｪｸﾄ
Dim TARGETOBJECT As Object

'カレントパス
Dim CuDir As String

'項目番号・断面番号・項目番号（下段）
Public kou_ID As Integer, dan_ID As Integer, kou_ID_D As Integer
Public kten_ID As Integer '経時変化図で出力する測点位置
Public muki_ID As Integer '断面分布図で出力する向き番号

'種類番号
Dim HENI As Integer

'描画変数
Dim PX As Single, PY As Single                         '座標位置
Dim PANK As String                                     '表示する文字列（Sub PPANK用）
Dim PANKSIZE As Integer                                'フォントサイズ（Sub AnkCsize用）
Dim ss As String                                       '表示する文字列（Sub KPUT用など）
Dim SIZ As Integer, XOFF As Integer, YOFF As Integer   'フォントサイズ・文字間隔（Sub KPUT用）
Dim PENC As Integer                                    '描画色（Sub PENJ用）
Dim PANKFM As String                                   '数値を描画するときのフォーマット文字列（Sub PANKFV用）
Dim PANKF As Variant                                   '表示する数値（Sub PANKFV用）
Dim SENC As Integer                                    '線種設定（Sub LTCD用）
Dim MKF As Integer                                     'マーカー設定（Sub MK用）
Dim RG As Integer, CSA As Integer, CEA As Integer      '円の描画（Sub CIL用）

'マーカー変数
Dim MKSW As Integer '0.描画しない 1.描画する
Dim MKBET As Single 'Ｘ座標係数
Dim MKSUU As Single '1ｸﾞﾘｯﾄﾞﾏｰｶｰ数

'パラメータ
Dim SD As Date, ED As Date         '経時変化図用 開始日・終了日
Dim XBUNKATU As Integer            '     〃     Ｘ軸分割数
Dim YBUNKATU As Integer            '     〃     Ｙ軸分割数
Dim Xtype As Integer               '     〃     Ｘ軸単位 1.日単位  2.分単位
Dim KDtype As Integer              '     〃     データ条件
Dim KTtype As Integer              '     〃     先頭時間
Dim Kanrisw As Integer             '     〃     管理値作図　0.Ｙｅｓ　1.Ｎｏ
Dim SEEKtime As Integer            '     〃     作図条件用 抽出時間
Dim SEEKMday(24) As String         '     〃         〃    １日でどの時間を作図するか時間を代入する
Dim HIZUKE(6) As Date              '断面分布図用 作図日時
Dim Xmin(2) As Single              '     〃     Ｘ軸最小値
Dim Xmax(2) As Single              '     〃     Ｘ軸最大値
Dim xBUN(2) As Single              '     〃     Ｘ軸分割数

'経時変化図用Ｙ軸パラメータ
Private Type Yjiku1
    kouNO As Integer
    danNO As Integer
    typeNO As Integer
    max As Single
    min As Single
    bunkatu As Integer
End Type
Dim Yjiku(5) As Yjiku1

'グラフ変数
Dim XGL As Single, YGL As Single   '起点座標
Dim x1 As Single, y1 As Single     '枠の始点(左下)座標
Dim XP As Single, YP As Single     '枠の長さ（0.1mm単位）
Dim YMAX As Single, YMIN As Single 'Ｙ軸最大値・Ｙ軸最小値
Dim YBAIRITU As Single             'Ｙ軸座標係数（１スケールの座標サイズ）
Dim YScl As Single                 '     〃    （１グリッドのスケール）
Dim YJIKUBAIRITU As Single         '     〃    （１グリッドの座標サイズ）
Dim KX1 As Single, KY1 As Single   '経時変化図用 グラフの始点(左下)座標
Dim KXP1 As Single, KYP1 As Single '     〃     グラフの長さ（0.1mm単位）
Dim maxD As Date                   '     〃     グラフ最終日
Dim XBAIRITU As Single             '     〃     Ｘ軸座標係数（１日の座標サイズ）
Dim MD As Double                   '     〃     開始日からの日数
Dim OLDAL As Double                '     〃     前回ﾏｰｶｰを描画したＸ軸座標
Dim SCL As Integer                 '断面分布図用 Ｘ軸の長さ（0.1mm単位）
Dim BunpuX1(2) As Integer          '     〃     Ｘ軸中心座標
Dim DBAIRITU(2) As Single          '     〃     Ｘ軸座標係数（１スケールの座標サイズ）
Dim DSC As Single                  '     〃     Ｘ軸座標係数（１グリッドのスケール）
Dim MaxLeng As Single              '     〃     Ｙ軸の最深部位置（ｍ単位）
Dim MinLeng As Single              '     〃     Ｙ軸の最上部位置（ｍ単位）
Dim BBB As Integer                 '     〃     作図日時番号
Dim DanmenSCL As Single            '     〃     1ｍの座標サイズ

Dim SX(4) As Single, SY(4) As Single, Spo(4) As Integer
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
    If Index = 1 Then VSPrinter1.Action = paChoosePrintPage
    If Index = 2 Then
        If muki_ID = 0 Then frmKeijiPara.Show Else frmBunpuPara.Show
    End If
End Sub

'**********************************************************************************************
'   フォームの初期設定
'   VSPrinter1の初期設定
'**********************************************************************************************
Private Sub Form_Load()

    Me.Height = Screen.Height - 420 '11000 '16590
    Me.Width = Screen.Width '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
'    Me.Height = 10500 '11000 '16590
'    Me.Width = 15500 '21210
'    Left = (Screen.Width - Me.Width) / 2
'    Top = 0
    
    Set TARGETOBJECT = VSPrinter1
        
    CuDir = CurDir
    
    BITMAP.Enabled = False
    
'    Me.Caption = "データシート"
'    Me.Height = 10840
'    Me.Width = 13785
'    Me.top = 0
'    Me.Left = Screen.Width - Width
    
    Screen.MousePointer = vbHourglass   '処理中は、マウスを砂時計にする
    
    '初期設定
    PrntDrvSW = True
    With VSPrinter1
        
        '印刷可能領域を表示
        .ShowGuides = gdHide
        
        .Preview = True
        '.Clear '.Cls
        .PenWidth = 8           ' 線幅
        'マウスをドラッグすることによりページのプレビューをスクロール
'        .MouseScroll = True
'        .MouseZoom = False
        '各ページの周りに描かれるページ枠を設定
        .PageBorder = pbAll
        'Printerコントロールの出力を全て画面へ
        .Preview = True
        'プレビュー画面の縮尺率
        .Zoom = 100 '80
        '用紙サイズをＡ４に設定
        If .PaperSizes(vbPRPSA4) = True Then
            .PaperSize = vbPRPSA4
        Else
            PrntDrvSW = False
            MsgBox "用紙サイズを設定できませんでした。", vbExclamation
            Screen.MousePointer = vbDefault   'マウスを既定値に戻す
            Exit Sub
        End If
        '用紙方向を横に設定
        .Orientation = orLandscape
        If .Error <> 0 Or .Orientation = orPortrait Then
            PrntDrvSW = False
            MsgBox "用紙方向を設定できませんでした。", vbExclamation
            Screen.MousePointer = vbDefault   'マウスを既定値に戻す
            Exit Sub
        End If
        
        .MarginLeft = 0 ' 1.5 * 567
        .MarginRight = 0 '1.5 * 567
        .MarginTop = 0 '3 * 567
        .MarginBottom = 0 '1.5 * 567
        .FontName = "ＭＳ ゴシック"
        .BrushStyle = bsTransparent
    End With
    
    BITMAP.Enabled = True
    
    Screen.MousePointer = vbDefault   'マウスを既定値に戻す
    Tyuudan = False

End Sub

Public Sub HeniBunpu(sclCK As Boolean)
    Dim i As Integer, j As Integer
    Dim FI1 As String
    
    Dim L As String, f As Integer
    Dim da As Date
    Dim pdSNG As Single, pdDBL As Double
    Dim piv As Long, Top As Long, po As Long
    Dim tdt(50) As Double
    Dim FLDno As Integer, FLDco As Integer
    Dim HI As Date
    Dim FileName As String
    
    Dim PSCL  As Single
    Dim INTLEN As Integer, TENLEN As Integer
    Dim YSCLMAX As Integer, YSCLMIN As Integer
    Dim ten As Integer
    Dim SW As Boolean, Dsw As Integer
    
    ZuOut = True
    INSATU.Enabled = False
    BITMAP.Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    
    mnuPrinterSet.Enabled = False
    
    Me.MousePointer = 11
    Me.WindowState = 0
    VSPrinter1.MarginLeft = 0 ' 1.5 * 567
    VSPrinter1.MarginRight = 0 '1.5 * 567
    VSPrinter1.MarginTop = 0 '3 * 567
    VSPrinter1.MarginBottom = 0 '1.5 * 567
    
    TARGETOBJECT.FontName = "ＭＳ ゴシック"
    TARGETOBJECT.BrushStyle = bsTransparent
    
    VSPrinter1.StartDoc
        
        x1 = 200: y1 = 100 + 20 '始点（左下）
        XP = 2550: YP = 1650         '長さ
        KY1 = 1600: KYP1 = -700
        If muki_ID = 1 Then
            KX1 = 0: KXP1 = 2550
        Else
            KX1 = 200: KXP1 = 2150
        End If
        
        PENC = 1: Call PENJ(TARGETOBJECT, PENC)
        SENC = 0: Call LTCD(TARGETOBJECT, SENC)
        Me.VSPrinter1.PenWidth = 8
        
        PX = x1 + KX1: PY = y1 + KY1
        If muki_ID = 1 Then FI1 = "図面3.emf"
        If muki_ID = 2 Then FI1 = "図面4.emf"
        TARGETOBJECT.x1 = 0: TARGETOBJECT.y1 = 0
        TARGETOBJECT.X2 = 0: TARGETOBJECT.Y2 = 0
        TARGETOBJECT.CalcPicture = LoadPicture(FI1)
        TARGETOBJECT.x1 = PX * 5.67
        TARGETOBJECT.y1 = TARGETOBJECT.PageHeight - PY * 5.67
        TARGETOBJECT.X2 = TARGETOBJECT.x1 + TARGETOBJECT.X2 '* 0.15
        TARGETOBJECT.Y2 = TARGETOBJECT.y1 + TARGETOBJECT.Y2 '* 0.15
        TARGETOBJECT.Picture = LoadPicture(FI1)
        
        kou_ID = 1: dan_ID = 1
        If muki_ID = 1 Then
            Spo(0) = 4
            SX(1) = 580: SY(1) = 0: Spo(1) = 1
            SX(2) = 920: SY(2) = -10: Spo(2) = 2
            SX(3) = 1470: SY(3) = -60: Spo(3) = 3
            SX(4) = 2080: SY(4) = -135: Spo(4) = 4
        Else
            Spo(0) = 3
            SX(1) = 400: SY(1) = -30: Spo(1) = 6
            SX(2) = 1100: SY(2) = 10: Spo(2) = 2
            SX(3) = 1800: SY(3) = -30: Spo(3) = 5
        End If
        
        'タイトル
        If muki_ID = 1 Then ss = Trim$(kou(kou_ID, 1).TI1) & "計測値（長辺方向）"
        If muki_ID = 2 Then ss = Trim$(kou(kou_ID, 1).TI1) & "計測値（短辺方向）"
        j = LenB(StrConv(ss, vbFromUnicode))
        PX = x1 + KX1 + KXP1 / 2 - (j / 2) * 25
        PY = y1 + 300: SIZ = 14: XOFF = 50: YOFF = 0
        Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
        
        '測定値
        FileName = KEISOKU.Data_path & DATA_DAT
        L = ""
        
        If Dir(FileName) = "" Then GoTo skip_2
        
        If sclCK = True Then
            f = FreeFile
            Open FileName For Input Shared As #f
                If LOF(f) - REC_LEN * 2 > 0 Then
                    Seek #f, LOF(f) - REC_LEN * 2
                End If
                Do While Not (EOF(f))
                    Line Input #f, L
                Loop
            Close #1
            If L = "" Then
                Hbunpu.SD = 0
            Else
                Hbunpu.SD = CDate(Mid$(L, 1, 19))
            End If
        Else
            Dsw = 0
            
            If SEEKmaster(Hbunpu.SD, piv, Top) = 0 Then
                GoTo skip_2
            Else
                po = piv
            End If
            
            f = FreeFile
            Open FileName For Input Access Read Shared As #f
            Seek #f, po
            Do While Not (EOF(f))
                Line Input #f, L
                If DateDiff("s", CDate(Left$(L, 19)), Hbunpu.SD) = 0 Then
                    Dsw = 1: Exit Do
                End If
            Loop
            Close (f)
            If Dsw = 0 Then L = ""
        
        End If
        
skip_2:
        
        YMAX = -999999: YMIN = 999999
        For i = 1 To Spo(0)
            FLDno = 20 + 10 * (Tbl(kou_ID, dan_ID, Spo(i)).FLD - 1): FLDco = 10
            If IsNumeric(Mid$(L, FLDno, FLDco)) = True Then
                pdDBL = CDbl(Mid$(L, FLDno, FLDco))
            Else
                pdDBL = 999999
            End If
            
            If Abs(pdDBL) >= 999999 Or Tbl(kou_ID, dan_ID, Spo(i)).Syo = 999999 Then
                tdt(Spo(i)) = 999999
            Else
                '2001/12/11
                tdt(Spo(i)) = (pdDBL - Tbl(kou_ID, dan_ID, Spo(i)).Syo) * Tbl(kou_ID, dan_ID, Spo(i)).Kei
                ''tdt(Spo(i)) = (-1) * (pdDBL - Tbl(kou_ID, dan_ID, Spo(i)).Syo) * Tbl(kou_ID, dan_ID, Spo(i)).Kei
            End If
            If YMIN > tdt(Spo(i)) Then YMIN = tdt(Spo(i))
            If YMAX < tdt(Spo(i)) Then YMAX = tdt(Spo(i))
        Next i
        
        
        '分布図軸
        If sclCK = True Then
            If YMAX - YMIN > 100 Then
                YMAX = Int(YMAX / 100) * 100 + 100
                YMIN = Int(YMIN / 100) * 100
                YBUNKATU = ((YMAX - YMIN) / 100) * 2
            ElseIf YMAX - YMIN > 10 Then
                YMAX = Int(YMAX / 10) * 10 + 10
                YMIN = Int(YMIN / 10) * 10
                YBUNKATU = (YMAX - YMIN) / 10
            Else
                YMAX = Int(YMAX) + 1
                YMIN = Int(YMIN) - 1
                YBUNKATU = YMAX - YMIN
            End If
            
            Hbunpu.YMIN = YMIN
            Hbunpu.YMAX = YMAX
            Hbunpu.yBUN = YBUNKATU
        Else
            YMIN = Hbunpu.YMIN
            YMAX = Hbunpu.YMAX
            YBUNKATU = Hbunpu.yBUN
        End If
        
        YScl = (YMAX - YMIN) / YBUNKATU
        YBAIRITU = KYP1 / (YMAX - YMIN) 'YBUNKATU / YScl
        YJIKUBAIRITU = KYP1 / YBUNKATU
        
        YSCLMAX = Len(Trim$(str$(Int(YMAX))))
        YSCLMIN = Len(Trim$(str$(Int(YMIN))))
        If YSCLMAX > YSCLMIN Then
            INTLEN = YSCLMAX
        Else
            INTLEN = YSCLMIN
        End If
        PANKFM = String$(INTLEN, "#")
        Mid$(PANKFM, Len(PANKFM), 1) = "0"
        
        ten = InStr(Trim$(str$(YScl)), ".")
        Select Case ten
            Case Is <> 0
                TENLEN = Len(Trim$(str$(YScl))) - ten
                PANKFM = PANKFM + "." + String$(TENLEN, "#")
                Mid$(PANKFM, Len(PANKFM), 1) = "0"
        End Select
        
        PX = x1 + KX1 + SX(1) - 200: PY = y1 + KY1 - 400: Call MMM(TARGETOBJECT, PX, PY)
        PX = x1 + KX1 + SX(1) - 200: PY = y1 + KY1 - 400 + KYP1: Call DDD(TARGETOBJECT, PX, PY)
        
        PX = x1 + KX1 + SX(Spo(0)) + 200: PY = y1 + KY1 - 400: Call MMM(TARGETOBJECT, PX, PY)
        PX = x1 + KX1 + SX(Spo(0)) + 200: PY = y1 + KY1 - 400 + KYP1: Call DDD(TARGETOBJECT, PX, PY)
        
        Me.VSPrinter1.PenColor = QBColor(7)
        SENC = 2: Call LTCD(TARGETOBJECT, SENC)
        Me.VSPrinter1.PenWidth = 0
        For i = 1 To Spo(0)
            PX = x1 + KX1 + SX(i): PY = y1 + KY1 - 400: Call MMM(TARGETOBJECT, PX, PY)
            PX = x1 + KX1 + SX(i): PY = y1 + KY1 - 400 + KYP1: Call DDD(TARGETOBJECT, PX, PY)
        Next i
        Me.VSPrinter1.PenWidth = 8
        SENC = 0: Call LTCD(TARGETOBJECT, SENC)
        Me.VSPrinter1.PenColor = QBColor(0)
        
        PSCL = YMIN
        For i = 0 To YBUNKATU
            If i = 0 Or i = YBUNKATU Then
                PX = x1 + KX1 + SX(1) - 200 - 15: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU: Call MMM(TARGETOBJECT, PX, PY)
                PX = x1 + KX1 + SX(Spo(0)) + 200: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU: Call DDD(TARGETOBJECT, PX, PY)
            Else
                PX = x1 + KX1 + SX(1) - 200 + 15: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU: Call MMM(TARGETOBJECT, PX, PY)
                PX = x1 + KX1 + SX(1) - 200 - 15: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU: Call DDD(TARGETOBJECT, PX, PY)
                Me.VSPrinter1.PenColor = QBColor(7)
                SENC = 2: Call LTCD(TARGETOBJECT, SENC)
                Me.VSPrinter1.PenWidth = 0
                PX = x1 + KX1 + SX(Spo(0)) + 200: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU: Call DDD(TARGETOBJECT, PX, PY)
                Me.VSPrinter1.PenWidth = 8
                SENC = 0: Call LTCD(TARGETOBJECT, SENC)
                Me.VSPrinter1.PenColor = QBColor(0)
            End If
            
            PANKSIZE = 10: Call AnkCsize(TARGETOBJECT, PANKSIZE)
            
            PX = x1 + KX1 + SX(1) - 200 - 150 - 10: PY = y1 + KY1 - 400 + i * YJIKUBAIRITU + 15
            PANKF = PSCL
            PANK = Format$(Format$(PSCL, PANKFM), "@@@@@@@@")
            Call PPANK(TARGETOBJECT, PANK, PX, PY)
            PSCL = PSCL + YScl
        Next i
        SW = False: MKF = 1
        For i = 1 To Spo(0)
            If tdt(Spo(i)) >= 999999 Then
                SW = False: GoTo skip_1
            Else
                PX = x1 + KX1 + SX(i)
                PY = y1 + KY1 - 400 + (tdt(Spo(i)) - YMIN) * YBAIRITU
            End If
            If SW = False Then
                Call MMM(TARGETOBJECT, PX, PY)
                SW = True
            Else
                Call DDD(TARGETOBJECT, PX, PY)
            End If
            Call MK(TARGETOBJECT, PX, PY, MKF)
skip_1:
        Next i
        
        
        ss = "(" & Trim$(kou(kou_ID, 1).Yu) & ")"
        PY = y1 + KY1 - 400 + 50: PX = x1 + KX1 + SX(1) - 200 - 80: SIZ = 10: XOFF = 38: YOFF = 0
        Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
        
        '計測値
        For i = 1 To Spo(0)
            PX = x1 + KX1 + SX(i): PY = y1 + KY1 + SY(i): Call MMM(TARGETOBJECT, PX, PY)
            PX = PX + 200: Call DDD(TARGETOBJECT, PX, PY)
            PY = PY + 70: Call DDD(TARGETOBJECT, PX, PY)
            PX = PX - 200: Call DDD(TARGETOBJECT, PX, PY)
            PY = PY - 70: Call DDD(TARGETOBJECT, PX, PY)
            
            If tdt(Spo(i)) >= 999999 Then
                ss = "  ******"
            Else
                ss = Format(Format(tdt(Spo(i)), "0.00"), "@@@@@@@@")    'SS = Format(Format(-123, "0.000"), "@@@@@@@@@@")
            End If
            PX = x1 + KX1 + SX(i) + 20: PY = y1 + KY1 + SY(i) + 55: SIZ = 12: XOFF = 42: YOFF = 0
            Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
            
            PX = x1 + KX1 + SX(i) + 220: PY = y1 + KY1 + SY(i) + 50: SIZ = 10: XOFF = 40: YOFF = 0
            ss = Trim(kou(kou_ID, 1).Yu)
            Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
        Next i
        
        '凡例
        PX = x1 + KX1 + 1700: PY = y1 + KY1 - 1200
        PENC = 1: Call PENJ(TARGETOBJECT, PENC)
        SIZ = 10: XOFF = 40: YOFF = 0
        If L = "" Then
            ss = "計測日時："
        Else
            ss = "計測日時：" & Format$(Hbunpu.SD, "ggge年m月d日 hh:nn:ss")
        End If
        Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
        
        If Tyuudan = True Then Exit Sub
    
    VSPrinter1.EndDoc
    
    Me.MousePointer = 0
    INSATU.Enabled = True
    BITMAP.Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    mnuPrinterSet.Enabled = True
    
    VSPrinter1.Visible = True
    cmdZoom(0).Visible = True: cmdZoom(1).Visible = True
    Label1(0).Visible = True
    
    ZuOut = False

    Label1(0).Caption = Format$(VSPrinter1.Zoom) & "%"

End Sub

Public Sub keijiPLot()
    Dim Xsw As Boolean, j As Integer
    Dim i As Integer
    Dim STtime As Date
    Dim KouCO As Integer, Hanco As Integer
    Dim ss As String, SS1 As String, SS2 As String, SS3 As String
    
    ZuOut = True
    INSATU.Enabled = False
    BITMAP.Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    mnuPrinterSet.Enabled = False
    
    Me.MousePointer = 11
    Me.WindowState = 0
    
    VSPrinter1.MarginLeft = 0 ' 1.5 * 567
    VSPrinter1.MarginRight = 0 '1.5 * 567
    VSPrinter1.MarginTop = 0 '3 * 567
    VSPrinter1.MarginBottom = 0 '1.5 * 567
    
    TARGETOBJECT.FontName = "ＭＳ ゴシック"
    TARGETOBJECT.BrushStyle = bsTransparent
    
    VSPrinter1.StartDoc
    
    
        PENC = 1: Call PENJ(TARGETOBJECT, PENC)
        SENC = 0: Call LTCD(TARGETOBJECT, SENC)
        
        x1 = 250: y1 = 150 + 20 '始点（左下）
        XP = 2500: YP = 1500         '長さ
        
        kou_ID = 1
        dan_ID = 1
        HENI = 1
        
'        SS = Trim(Tbl(kou_ID, dan_ID, kten_ID).HAN)
'        Me.Caption = SS & "経時変化図"
'        Me.Refresh
        
        SD = Hkeiji.SD
        ED = Hkeiji.ED
        If DateDiff("n", SD, ED) > 1460 Then Xtype = 1 Else Xtype = 2
        XBUNKATU = Hkeiji.xBUN
        YMIN = Hkeiji.YMIN
        YMAX = Hkeiji.YMAX
        YBUNKATU = Hkeiji.yBUN
        
        MKSUU = 50
        KDtype = 0
        KTtype = 0
        
        
        KY1 = 1500: KYP1 = -1500
        KX1 = 0: KXP1 = 2500
        PENC = 1: Call PENJ(TARGETOBJECT, PENC)
        SENC = 0: Call LTCD(TARGETOBJECT, SENC)
        
        'タイトル
        ss = Trim$(kou(kou_ID, 1).TI1) & "経時変化図（" & StrConv(Trim(Tbl(kou_ID, dan_ID, kten_ID).HAN), vbWide) & "）"
        j = LenB(StrConv(ss, vbFromUnicode))
        PX = x1 + KX1 + KXP1 / 2 - (j / 2) * 22
        PY = y1 + KY1 + 130: SIZ = 12: XOFF = 44: YOFF = 0
        Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
        
        Xsw = True
        Call KFLAME(Xsw)
        Call KPLOT1
        If Tyuudan = True Then Exit Sub
    
    VSPrinter1.EndDoc
    
    Me.MousePointer = 0
    INSATU.Enabled = True
    BITMAP.Enabled = True
    Command1(1).Enabled = True
    Command1(2).Enabled = True
    mnuPrinterSet.Enabled = True
    
    VSPrinter1.Visible = True
    cmdZoom(0).Visible = True: cmdZoom(1).Visible = True
    Label1(0).Visible = True
    
    ZuOut = False

    Label1(0).Caption = Format$(VSPrinter1.Zoom) & "%"

End Sub

'**********************************************************************************************
'   経時変化図 グリッド線・目盛り描画
'**********************************************************************************************
Private Sub KFLAME(Xsw As Boolean)
    Dim ZENNISSUU As Long
    Dim NENDO As Variant
    Dim PSCL As Single
    Dim OLDPX As Single
    Dim OLDNENDO As Variant
    Dim INTLEN As Integer, TENLEN As Integer
    Dim YSCLMAX As Integer, YSCLMIN As Integer
    Dim ten As Integer
    Dim DKANKAKU As Integer
    Dim NISSUU As Integer
    Dim SX As Single, SY As Single
    Dim i As Integer
    Dim ssP As String, ssM As String
    
    If Xtype = 1 Then
        ''ZENNISSUU = DateDiff("y", sd, ed) + 1
        If (ED - SD) = Int(ED - SD) Then
            ZENNISSUU = ED - SD
        Else
            ZENNISSUU = Int(ED - SD) + 1
        End If
        If (ZENNISSUU / XBUNKATU) = Int(ZENNISSUU / XBUNKATU) Then
            DKANKAKU = Int(ZENNISSUU / XBUNKATU)
        Else
            DKANKAKU = Int(ZENNISSUU / XBUNKATU) + 1
        End If
        XBAIRITU = KXP1 / (DKANKAKU * XBUNKATU)  'ZENNISSUU
        'ed = DateAdd("y", DKANKAKU * XBUNKATU, sd)
        
        If Xsw = True Then
            ss = "日付": PX = x1 + KX1 - 150: PY = y1 + KY1 + KYP1 - 50: SIZ = 9: XOFF = 40: YOFF = 0
            Call KPUT(TARGETOBJECT, ss, PX, PY + 26, SIZ, XOFF, YOFF, 0)
'            SS = "日数": PX = X1 + KX1 - 150: PY = Y1 + KY1 + KYP1 - 50: SIZ = 9: XOFF = 40: YOFF = 0
'            Call KPUT(TARGETOBJECT, SS, PX, PY + 26, SIZ, XOFF, YOFF, 0)
'
'            SS = "日付": PX = X1 + KX1 - 150: PY = Y1 + KY1 + KYP1 - 90: SIZ = 9: XOFF = 40: YOFF = 0
'            Call KPUT(TARGETOBJECT, SS, PX, PY + 26, SIZ, XOFF, YOFF, 0)
'
'            'If XBUNKATU > 7 Then
'            If DatePart("yyyy", DateAdd("y", XBUNKATU * DKANKAKU, sd)) <> DatePart("yyyy", sd) Then
'                PY = Y1 + KY1 + KYP1 - 150
'            Else
'                PY = Y1 + KY1 + KYP1 - 130
'            End If
'            SS = "時刻": PX = X1 + KX1 - 150:  SIZ = 9: XOFF = 40: YOFF = 0
'            Call KPUT(TARGETOBJECT, SS, PX, PY + 26, SIZ, XOFF, YOFF, 0)
        End If
        
        OLDNENDO = 0
        For i = 0 To XBUNKATU
            NENDO = DateAdd("y", i * DKANKAKU, SD)
            NISSUU = DateDiff("y", SD, NENDO)
         
            PX = x1 + KX1 + NISSUU * XBAIRITU: PY = y1 + KY1
            Call MMM(TARGETOBJECT, PX, PY)
            PX = x1 + KX1 + NISSUU * XBAIRITU: PY = y1 + KY1 + KYP1
            Call DDD(TARGETOBJECT, PX, PY)
            
            If Xsw = True Then

                'PX = X1 + KX2 + NISSUU * XBAIRITU: PY = Y1 + KY2 + KYP2
                PANKSIZE = 8:  Call AnkCsize(TARGETOBJECT, PANKSIZE)
                PX = PX - 10
                
'                PY = Y1 + KY1 + KYP1 - 20: PANK = CStr(NISSUU)
'                Call PPANK(TARGETOBJECT, PANK, PX, PY)
                
                'If XBUNKATU > 7 Then
'                If DatePart("yyyy", DateAdd("y", XBUNKATU * DKANKAKU, sd)) <> DatePart("yyyy", sd) Then
'                    If Format$(OLDNENDO, "YYYY") <> Format$(NENDO, "YYYY") Then
'                        PY = Y1 + KY1 + KYP1 - 60: PANK = "[" & Format$(NENDO, "YYYY") & "年]"
'                        Call PPANK(TARGETOBJECT, PANK, PX - 40, PY)
'                    End If
'
'                    PY = Y1 + KY1 + KYP1 - 100: PANK = Format$(NENDO, "M/D")
'                    Call PPANK(TARGETOBJECT, PANK, PX, PY)
'
'                    If OLDNENDO = 0 Or Format$(OLDNENDO, "h:nn") <> Format$(NENDO, "h:nn") Then
'                        PY = Y1 + KY1 + KYP1 - 140: PANK = Format$(NENDO, "h:nn")
'                        Call PPANK(TARGETOBJECT, PANK, PX, PY)
'                    End If
'                Else
                    PY = y1 + KY1 + KYP1 - 20: PANK = Format$(NENDO, "yy/M/D")
                    Call PPANK(TARGETOBJECT, PANK, PX, PY)
'                    If OLDNENDO = 0 Or Format$(OLDNENDO, "h:nn") <> Format$(NENDO, "h:nn") Then
'                        PY = Y1 + KY1 + KYP1 - 110: PANK = Format$(NENDO, "h:nn")
'                        Call PPANK(TARGETOBJECT, PANK, PX, PY)
'                    End If
'                End If
            End If
            
            OLDPX = x1 + KX1 + NISSUU * XBAIRITU
            OLDNENDO = NENDO
        Next i
        maxD = NENDO
    Else
        ZENNISSUU = DateDiff("n", SD, ED)
        If (ZENNISSUU / XBUNKATU) = Int(ZENNISSUU / XBUNKATU) Then
            DKANKAKU = Int(ZENNISSUU / XBUNKATU)
        Else
            DKANKAKU = Int(ZENNISSUU / XBUNKATU) + 1
        End If
        XBAIRITU = KXP1 / (DKANKAKU * XBUNKATU)  'ZENNISSUU
        XBAIRITU = KXP1 / (DKANKAKU * XBUNKATU)  'ZENNISSUU
        'ed = DateAdd("n", DKANKAKU * XBUNKATU, sd)
        
        If Xsw = True Then
            ss = "経過時間（分）": PX = x1 + KX1 - 220: PY = y1 + KY1 + KYP1 - 50: SIZ = 9: XOFF = 30: YOFF = 0
            Call KPUT(TARGETOBJECT, ss, PX, PY + 26, SIZ, XOFF, YOFF, 0)
            ss = "時刻": PX = x1 + KX1 - 170: PY = y1 + KY1 + KYP1 - 90: SIZ = 9: XOFF = 40: YOFF = 0
            Call KPUT(TARGETOBJECT, ss, PX, PY + 26, SIZ, XOFF, YOFF, 0)
            ss = "日付": PX = x1 + KX1 - 170: PY = y1 + KY1 + KYP1 - 130: SIZ = 9: XOFF = 40: YOFF = 0
            Call KPUT(TARGETOBJECT, ss, PX, PY + 26, SIZ, XOFF, YOFF, 0)
        End If
        
        OLDNENDO = 0
        For i = 0 To XBUNKATU
            NENDO = DateAdd("n", i * DKANKAKU, SD)
            NISSUU = DateDiff("n", SD, NENDO)
         
            PX = x1 + KX1 + NISSUU * XBAIRITU: PY = y1 + KY1
            Call MMM(TARGETOBJECT, PX, PY)
            PX = x1 + KX1 + NISSUU * XBAIRITU: PY = y1 + KY1 + KYP1
            Call DDD(TARGETOBJECT, PX, PY)
            
            'PX = X1 + KX2 + NISSUU * XBAIRITU: PY = Y1 + KY2 + KYP2
            If Xsw = True Then

                PANKSIZE = 8: Call AnkCsize(TARGETOBJECT, PANKSIZE)
                PX = PX - 10
                PY = y1 + KY1 + KYP1 - 20: PANK = CStr(NISSUU)
                Call PPANK(TARGETOBJECT, PANK, PX, PY)
    
                PY = y1 + KY1 + KYP1 - 60: PANK = Format$(NENDO, "h:nn")
                Call PPANK(TARGETOBJECT, PANK, PX, PY)
    
                If Format$(OLDNENDO, "YYYY/m/d") <> Format$(NENDO, "YYYY/m/d") Then
                    PY = y1 + KY1 + KYP1 - 100: PANK = Format$(NENDO, "YYYY/M/D")
                    Call PPANK(TARGETOBJECT, PANK, PX, PY)
                End If
            End If
            
            OLDPX = x1 + KX1 + NISSUU * XBAIRITU
            OLDNENDO = NENDO
        Next i
        maxD = NENDO
    End If
    
    
    ''XJIKUBAIRITU = KXP1 / XBUNKATU
    'MKBET = DKANKAKU / FrmParameta.Text03(1).Text                  'ﾏｰｶｰ間隔（日）
    If MKSUU = 0 Then
        MKSW = 0
    Else
        MKSW = 1
        MKBET = DKANKAKU / MKSUU
    End If
    
    'Ｙ軸
    YScl = (YMAX - YMIN) / YBUNKATU
    YBAIRITU = KYP1 / (YMAX - YMIN) 'YBUNKATU / YScl
    YJIKUBAIRITU = KYP1 / YBUNKATU
    
    YSCLMAX = Len(Trim$(str$(Int(YMAX))))
    YSCLMIN = Len(Trim$(str$(Int(YMIN))))
    If YSCLMAX > YSCLMIN Then
        INTLEN = YSCLMAX
    Else
        INTLEN = YSCLMIN
    End If
    PANKFM = String$(INTLEN, "#")
    Mid$(PANKFM, Len(PANKFM), 1) = "0"
    
    ten = InStr(Trim$(str$(YScl)), ".")
    Select Case ten
        Case Is <> 0
            TENLEN = Len(Trim$(str$(YScl))) - ten
            PANKFM = PANKFM + "." + String$(TENLEN, "#")
            Mid$(PANKFM, Len(PANKFM), 1) = "0"
    End Select
    
    PSCL = YMIN
    For i = 0 To YBUNKATU
        PX = x1 + KX1: PY = y1 + KY1 + i * YJIKUBAIRITU
        Call MMM(TARGETOBJECT, PX, PY)
        PX = x1 + KX1 + KXP1: PY = y1 + KY1 + i * YJIKUBAIRITU
        Call DDD(TARGETOBJECT, PX, PY)
        
        PANKSIZE = 8: Call AnkCsize(TARGETOBJECT, PANKSIZE)
        
        PX = x1 + KX1 - 130: PY = y1 + KY1 + i * YJIKUBAIRITU + 15
        PANKF = PSCL
        PANK = Format$(Format$(PSCL, PANKFM), "@@@@@@@@")
        Call PPANK(TARGETOBJECT, PANK, PX, PY)
        PSCL = PSCL + YScl
    Next i
    
    PX = x1 + KX1 - 170: PY = y1 + KY1 + KYP1 / 2 + 50: SIZ = 10: XOFF = 0: YOFF = -40
    ss = Trim$(kou(kou_ID, HENI).Yt)
    Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
    
    PY = y1 + KY1 + KYP1 / 2 - Len(ss) * 20 - 40
    ss = "(" & Trim$(kou(kou_ID, HENI).Yu) & ")"
    PX = x1 + KX1 - 150 - Len(ss) * 10: SIZ = 7: XOFF = 40: YOFF = 0
    Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
    

End Sub

'**********************************************************************************************
'   経時変化図 データ描画
'**********************************************************************************************
Private Sub KPLOT1()
    Dim i As Integer, j As Integer
    Dim L As String
    Dim da As Date
    Dim pdSNG As Single, pdDBL As Double
    Dim po  As Long
    Dim tdt(50) As Double
    Dim FLDno As Integer, FLDco As Integer
    Dim HI As Date
    Dim sen As Integer
    Dim FileName As String
    
    Dim SW(50) As Integer
    Dim OLDPD(50) As Double
    Dim OLDmd(50) As Double
    
    OLDAL = -99: 'ﾏｰｶｰ間隔初期値
    For i = 1 To Tbl(kou_ID, dan_ID, 0).ten: SW(i) = 0: Next i
    
    FileName = KEISOKU.Data_path & DATA_DAT
    
    If Dir(FileName) = "" Then
        Open FileName For Output As #1
        Close #1
    End If
    
    po = STARTpoint(SD)
    
    Open FileName For Input Shared As #1
    Seek #1, po
    Do While Not (EOF(1))

DoEvents
If Tyuudan = True Then Exit Do
        
        Line Input #1, L
        da = CDate(Mid$(L, 1, 19))
        If da < SD Then GoTo Kskip
        If da > ED Then Exit Do
        
        If 0 < KDtype Then
            For i = 1 To 24 / SEEKtime
                If Format(da, "hh:nn:ss") = SEEKMday(i) Then Exit For
            Next i
            If i > (24 / SEEKtime) Then GoTo Kskip
        End If
        
        HI = da
        If Xtype = 1 Then
            MD = DateDiff("s", SD, HI) / 86400#
        Else
            MD = DateDiff("s", SD, HI) / 60#
        End If
        
        FLDno = 20 + 10 * (Tbl(kou_ID, dan_ID, kten_ID).FLD - 1): FLDco = 10
        If IsNumeric(Mid$(L, FLDno, FLDco)) = True Then
            pdDBL = CDbl(Mid$(L, FLDno, FLDco))
        Else
            pdDBL = 999999
        End If
        
        If Abs(pdDBL) >= 999999 Or Tbl(kou_ID, dan_ID, kten_ID).Syo = 999999 Then
            tdt(kten_ID) = 999999
        Else
            '2001/12/11
            tdt(kten_ID) = (pdDBL - Tbl(kou_ID, dan_ID, kten_ID).Syo) * Tbl(kou_ID, dan_ID, kten_ID).Kei
''            tdt(kten_ID) = (-1) * (pdDBL - Tbl(kou_ID, dan_ID, kten_ID).Syo) * Tbl(kou_ID, dan_ID, kten_ID).Kei
        End If
'''        For i = 1 To Tbl(kou_ID, dan_ID, 0).ten
'''            FLDno = 20 + 8 * (Tbl(kou_ID, dan_ID, i).FLD - 1): FLDco = 8
'''            If IsNumeric(Mid$(L, FLDno, FLDco)) = True Then
'''                pdSNG = CSng(Mid$(L, FLDno, FLDco))
'''            Else
'''                pdSNG = 999999
'''            End If
'''
'''            If Abs(pdSNG) >= 999999 Or Tbl(kou_ID, dan_ID, i).Syo = 999999 Then
'''                tdt(i) = 999999
'''            Else
'''                If kou_ID = 2 Then Call SinsyukuCAL(dan_ID, i, da, pdSNG)
'''                tdt(i) = (pdSNG - Tbl(kou_ID, dan_ID, i).Syo) * Tbl(kou_ID, dan_ID, i).Kei
'''            End If
'''        Next i
'''        Call KEISAN(kou_ID, dan_ID, HENI, tdt())
        
        
        sen = 1
        pdDBL = tdt(kten_ID)
        
        sen = sen + 1
        PENC = sen: Call PENJ(TARGETOBJECT, PENC)
        MKF = sen - 1
        If sen > 7 Then SENC = 2 Else SENC = 0
        Call LTCD(TARGETOBJECT, SENC)
        
        If Abs(pdDBL) = 999999 Then GoTo Kskip2
        
        Select Case SW(i)
            Case 0
                PX = x1 + KX1 + MD * XBAIRITU
                PY = y1 + KY1 + (pdDBL - YMIN) * YBAIRITU
                Call MMM(TARGETOBJECT, PX, PY)
                SW(i) = 1
            Case 1
                PX = x1 + KX1 + OLDmd(i) * XBAIRITU
                PY = y1 + KY1 + (OLDPD(i) - YMIN) * YBAIRITU
                Call MMM(TARGETOBJECT, PX, PY)
                PX = x1 + KX1 + MD * XBAIRITU
                PY = y1 + KY1 + (pdDBL - YMIN) * YBAIRITU
                Call DDD(TARGETOBJECT, PX, PY)
        End Select

        If MKSW = 1 And (MD - OLDAL) >= MKBET Then
            Call LTCD(TARGETOBJECT, 0)
            Call MK(TARGETOBJECT, PX, PY, MKF)
            Call LTCD(TARGETOBJECT, SENC)
            'If i = Tbl(dan_ID, kou_ID, 0).FLD Then OLDAL = MD
        End If
        
        OLDPD(i) = pdDBL
        OLDmd(i) = MD
Kskip2:
        
        If MKSW = 1 And (MD - OLDAL) >= MKBET Then OLDAL = MD
Kskip:
    Loop
    Close
'Debug.Print HI, tdt(BBB)

    '現在値
    PENC = 1: Call PENJ(TARGETOBJECT, PENC)
    ss = "現在値："
    PX = x1 + KX1 + 1800 - 36 * 4: PY = y1 + KY1 + 50: SIZ = 10: XOFF = 36: YOFF = 0
    Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
    
    If tdt(kten_ID) >= 999999 Then
        ss = "    ******"
    Else
        ss = Format(Format(tdt(kten_ID), "0.00"), "@@@@@@@@@@")    'SS = Format(Format(-123, "0.000"), "@@@@@@@@@@")
    End If
    PX = x1 + KX1 + 1800: PY = y1 + KY1 + 50: SIZ = 10: XOFF = 36: YOFF = 0
    Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
    
    PX = x1 + KX1 + 1800 + 220: PY = y1 + KY1 + 50: SIZ = 9: XOFF = 36: YOFF = 0
    ss = Trim(kou(kou_ID, 1).Yu)
    Call KPUT(TARGETOBJECT, ss, PX, PY, SIZ, XOFF, YOFF, 0)
End Sub

'**********************************************************************************************
'   作業中に中断した場合は、印刷ジョブを削除します。
'**********************************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ZuOut = True Then Tyuudan = True: VSPrinter1.KillDoc
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
'    With Label1(0)
'        .Left = cmdZoom(0).Left + cmdZoom(0).Width + mintControlMargin + cmdZoom(1).Width + mintControlMargin
'        .top = 175 'Me.ScaleHeight - mintControlMargin - cmdZoom(1).Height + 100
'    End With
'    With Command1
'        .Left = Me.ScaleWidth - 1680
'        .top = 75
'    End With
'
'    With VSPrinter1
'        .Left = mintControlMargin + 1500
'        .top = 480 'mintControlMargin
'        .Height = Me.ScaleHeight - cmdZoom(0).Height - 3 * mintControlMargin
'        .Width = Me.ScaleWidth - 2 * mintControlMargin - 1500
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
    With Frame1
        .Left = Me.ScaleWidth - 1500
        .Top = 2640 'Me.ScaleHeight - 2135
    End With

    With VSPrinter1
        .Left = mintControlMargin
        .Top = mintControlMargin
        .Height = Me.ScaleHeight - 2 * mintControlMargin
        .Width = Me.ScaleWidth - 2 * mintControlMargin - 1500
    End With
'Debug.Print Me.Width, Me.Height
End Sub

Private Sub INSATU_Click()
    If ZuOut = True Then Exit Sub
    VSPrinter1.Action = paChoosePrintPage
End Sub

Private Sub mnuPrinterSet_Click()
    VSPrinter1.Action = paChoosePrinter
End Sub

Private Sub SYUURYOU_Click()
    Unload Me
End Sub

Private Sub BITMAP_Click()
    If ZuOut = True Then Exit Sub
'    CommonDialog1.DefaultExt = ".BMP"
'    CommonDialog1.Filter = "ビットマップ ファイル (*.BMP)|*.BMP"
'    CommonDialog1.Flags = &H2& Or &H8&
'
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'    CommonDialog1.ShowSave
'    On Error GoTo 0
'
'
'    frmBitmapSave.Picture1.Picture = Me.VSPrinter1.Picture
'    SavePicture frmBitmapSave.Picture1.Image, CommonDialog1.FileName
'
'    ChDrive CuDir
'    ChDir CuDir
'
'    Exit Sub
'
'ErrHandler:
'    ' ユーザーが [キャンセル] をクリックしました。
'    ChDrive CuDir
'    ChDir CuDir
'    Exit Sub
End Sub


