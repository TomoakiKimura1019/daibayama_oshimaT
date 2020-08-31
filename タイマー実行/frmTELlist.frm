VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "VSFLEX8L.OCX"



Begin VB.Form frmTELlist 
   Caption         =   "緊急時連絡先"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12825
   Icon            =   "frmTELlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12825
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   12000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "印刷"
      Height          =   495
      Left            =   11280
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "メニューに戻る"
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VSFlex8LCtl.VSFlexGrid MSFlexGrid1 
      Height          =   8865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   15637
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedRows       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
      NavBar = 0
      Height          =   495
      Left            =   11400
      TabIndex        =   2
      Top             =   1920
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
      _ConvInfo       =   1
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
      Zoom            =   -0.377833753148615
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   5
      MouseZoom       =   2
      MouseScroll     =   -1  'True
      MousePage       =   -1  'True
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
      HTMLStyle       =   1
   End
End
Attribute VB_Name = "frmTELlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'緊急時連絡先
Private Type telList1
    no As Integer
    Office1 As String * 30
    Office2 As String * 30
    Office3 As String * 30
    Tel1 As String * 20
    Tel2 As String * 20
    Name1 As String * 16
    Name2 As String * 16
End Type
Dim telList(10) As telList1

Sub telread()
    Dim f As Integer, i As Integer, L As String
    i = 0
    f = FreeFile
    Open KEISOKU.Tabl_path & "緊急連絡.dat" For Input Shared As #f
        Do While Not (EOF(f))
            Line Input #f, L
            If Left$(L, 1) <> ":" Then
                i = i + 1
                telList(i).no = i
                telList(i).Office1 = Trim(SEEKmoji(L, 5, 30))
                telList(i).Office2 = Trim(SEEKmoji(L, 35, 30))
                telList(i).Office3 = Trim(SEEKmoji(L, 65, 30))
                telList(i).Tel1 = Trim(SEEKmoji(L, 95, 20))
                telList(i).Tel2 = Trim(SEEKmoji(L, 115, 20))
                telList(i).Name1 = Trim(SEEKmoji(L, 135, 16))
                telList(i).Name2 = Trim(SEEKmoji(L, 151, 16))
            End If
        Loop
    Close #f
    telList(0).no = i
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
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

Private Sub Form_Load()
    Dim i As Integer
    Dim SS As String
    
    Me.Height = Screen.Height - 420 '11000 '16590
    Me.Width = Screen.Width '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
'    Me.Height = 9500 '16590
'    Me.Width = 13500 '21210
'    Left = (Screen.Width - Me.Width) / 2
'    Top = 0

    Call telread

    If telList(0).no < 5 Then
        MSFlexGrid1.Rows = 7
    Else
        MSFlexGrid1.Rows = telList(0).no + 2
    End If
    MSFlexGrid1.Width = 10860
    MSFlexGrid1.Height = 8150
    
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.ColWidth(1) = 3700
    MSFlexGrid1.ColWidth(2) = 2500
    MSFlexGrid1.ColWidth(3) = 2030
    MSFlexGrid1.ColWidth(4) = 2030

    MSFlexGrid1.RowHeight(-1) = 750
    MSFlexGrid1.RowHeight(0) = 250
    MSFlexGrid1.RowHeight(1) = 250
    
    MSFlexGrid1.ColAlignment(0) = 4
    For i = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.ColAlignment(i) = 1
        MSFlexGrid1.Row = 0: MSFlexGrid1.Col = i: MSFlexGrid1.CellAlignment = 4
        MSFlexGrid1.Row = 1: MSFlexGrid1.Col = i: MSFlexGrid1.CellAlignment = 4
    Next i
    
    MSFlexGrid1.MergeCells = 1
    MSFlexGrid1.MergeRow(0) = True
    MSFlexGrid1.MergeRow(1) = True
    MSFlexGrid1.MergeCol(0) = True
    MSFlexGrid1.MergeCol(1) = True
    MSFlexGrid1.MergeCol(2) = True
        
    MSFlexGrid1.TextMatrix(0, 0) = " "
    MSFlexGrid1.TextMatrix(1, 0) = " "
    MSFlexGrid1.TextMatrix(0, 1) = "会社名"
    MSFlexGrid1.TextMatrix(1, 1) = "会社名"
    MSFlexGrid1.TextMatrix(0, 2) = "電話"
    MSFlexGrid1.TextMatrix(1, 2) = "電話"
    MSFlexGrid1.TextMatrix(0, 3) = "管理者"
    MSFlexGrid1.TextMatrix(1, 3) = "正"
    MSFlexGrid1.TextMatrix(0, 4) = "管理者"
    MSFlexGrid1.TextMatrix(1, 4) = "副"
    MSFlexGrid1.WordWrap = True
    
    MSFlexGrid1.MergeCells = 1
    
    For i = 1 To telList(0).no
        MSFlexGrid1.TextMatrix(i + 1, 0) = telList(i).no
        SS = Trim(telList(i).Office1)
        If Trim(telList(i).Office2) <> "" Then SS = SS & Chr(13) & Trim(telList(i).Office2)
        If Trim(telList(i).Office3) <> "" Then SS = SS & Chr(13) & Trim(telList(i).Office3)
        MSFlexGrid1.TextMatrix(i + 1, 1) = SS
        
        SS = Trim(telList(i).Tel1)
        If Trim(telList(i).Tel2) <> "" Then SS = SS & Chr(13) & Trim(telList(i).Tel2)
        MSFlexGrid1.TextMatrix(i + 1, 2) = SS
        MSFlexGrid1.TextMatrix(i + 1, 3) = Trim(telList(i).Name1)
        MSFlexGrid1.TextMatrix(i + 1, 4) = Trim(telList(i).Name2)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMenu.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With MSFlexGrid1
        .Height = Me.Height - 630
        .Width = Me.Width - 2085
    End With
    Command1.Left = Me.Width - 1665
    Command2.Left = Me.Width - 1665
End Sub


