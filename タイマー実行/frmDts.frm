VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "VSPrint8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDts 
   Caption         =   "�f�[�^�V�[�g"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11175
   Icon            =   "frmDts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11175
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   495
      Index           =   1
      Left            =   9360
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�o�͓��ύX"
      Height          =   495
      Index           =   2
      Left            =   9360
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "�y�[�W"
      Height          =   1695
      Index           =   1
      Left            =   8880
      TabIndex        =   7
      Top             =   2520
      Width           =   2055
      Begin VB.CommandButton cmdPage 
         Caption         =   "��擪"
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
         Caption         =   "�ŏI��"
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
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
         BackStyle       =   1  '�s����
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Y�[��"
      Height          =   1815
      Index           =   0
      Left            =   8880
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton cmdZoom 
         Caption         =   "�k��"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdZoom 
         Caption         =   "�g��"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��������
         BackStyle       =   0  '����
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
         FillStyle       =   0  '�h��Ԃ�
         Height          =   285
         Index           =   0
         Left            =   1080
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "�\���{��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���j���[�ɖ߂�"
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
         Name            =   "�l�r �S�V�b�N"
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
      AbortCaption    =   "�����..."
      AbortTextButton =   "��ݾ�"
      AbortTextDevice =   "�o�͐� %s(%s)"
      AbortTextPage   =   "%d �߰�ޖڂ������"
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
      NavBarMenuText  =   "�߰�ޑS��(&P)|�߰�ޕ�(&W)|2�߰��(&T)|��Ȳ�(&N)"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��"
      Visible         =   0   'False
      Begin VB.Menu mnuPrinterSet 
         Caption         =   "�������ݒ�"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "���"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "̧�ق֕ۑ�"
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "�I��"
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
'   �f�[�^�V�[�g��\���E����E�t�@�C���ۑ�����
'
'       ����ݒ� �T�C�Y���`�S
'                ��������
'                �����t�H���g���l�r ����
'                �]������25mm�A��20mm�A��10mm�A�E10mm
'----------------------------------------------------------------------------------------------
'�o�͐� 1.�\��  2.�t�@�C���ۑ�
Dim PrnMode As Integer

'�\����Ɗm�F True=��ƒ� False=��Ɗ���
Public HYOUsw As Boolean

'True=��ƒ��ɒ��f���ꂽ�ꍇ
Dim Tyuudan As Boolean

'���ڔԍ��E�f�ʔԍ�
Dim kou_ID As Integer, dan_ID As Integer

'��ޔԍ�
Dim HENI As Integer

'�o�͐��޼ު��
Private TARGETOBJECT As Object

'�f�[�^�V�[�g�s������
Dim strBody  As String

'�J�����g�p�X
Dim CuDir As String

'�p�����[�^
Dim DS_SD As Date, DS_ED As Date '�J�n���E�I����
Dim SDtype As Integer            '�f�[�^����
Dim STtype As Integer            '�擪����
Dim SEEKtime As Integer          '��\�����p ���o����
Dim SEEKMday(24) As String       '    �V    �P���łǂ̎��Ԃ���}���邩���Ԃ�������
Dim sTYPE As Integer             '�\���`�� 0.������ 1.����l
Dim DTS_Col_WIDTH As Integer     '��̕�
Dim DTS_Col_MAX   As Integer     '��

'**********************************************************************************************
'   ��\�����i��ʕ`��j
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
    
    '����J�n
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
    
    If PrnMode = 1 Then VSPrinter1.EndDoc             ' ������I�����܂��B
    
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
'   �\���y�[�W�ړ�
'       Index 0=�擪 1=�ŏI
'**********************************************************************************************
Private Sub cmdPage_Click(Index As Integer)

    If Index = 0 Then
        VSPrinter1.PreviewPage = 1
        scrlPage.Value = 1
    Else
        VSPrinter1.PreviewPage = VSPrinter1.PageCount
        scrlPage.Value = VSPrinter1.PageCount
    End If
    
    Label1(1).Caption = Format$(VSPrinter1.PreviewPage) & "/" & Format$(VSPrinter1.PageCount) & " �߰��"

End Sub

'**********************************************************************************************
'   �\������Ă�f�[�^�V�[�g�̏k�ڂ�ݒ肵�܂��B
'       �k�ڲ�����فFZoomParam
'       �ő�k�ڗ��F150
'       �ŏ��k�ڗ��F20
'**********************************************************************************************
Private Sub cmdZoom_Click(Index As Integer)
    Const ZoomParam = 10
    
    With VSPrinter1
        Select Case Index
            Case 0
                .Zoom = .Zoom + ZoomParam
                'Zoom�̋��e�͈͂��O���ꍇ�A[�g��]�{�^�����g�p�s�\�ɐݒ�
                If .Zoom > 150 - ZoomParam Then
                    cmdZoom(0).Enabled = False
                End If
                '[�k��]�{�^�����g�p�\�ɐݒ�
                If Not cmdZoom(1).Enabled Then
                    cmdZoom(1).Enabled = True
                End If
            Case 1
                .Zoom = .Zoom - ZoomParam
                'Zoom�̋��e�͈͂��O���ꍇ�A[�k��]�{�^�����g�p�s�\�ɐݒ�
                If .Zoom <= 0 + ZoomParam Then
                    cmdZoom(1).Enabled = False
                End If
                '[�g��]�{�^�����g�p�\�ɐݒ�
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
'   �t�H�[���̏����ݒ�
'   VSPrinter1�̏����ݒ�
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
    
'    Me.Caption = "�f�[�^�V�[�g"
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
        
        '����\�̈��\��
        .ShowGuides = gdShow: 'Me.mnuSGuides.Checked = True
        '�}�E�X���h���b�O���邱�Ƃɂ��y�[�W�̃v���r���[���X�N���[��
'        .MouseScroll = False
'        .MouseZoom = False
        
        '�e�y�[�W�̎���ɕ`�����y�[�W�g��ݒ�
        .PageBorder = 0  'pbAll
        'Printer�R���g���[���̏o�͂�S�ĉ�ʂ�
        .Preview = True
        '�v���r���[��ʂ̏k�ڗ�
        .Zoom = 100
        .ZoomMode = zmPercentage
        '�p���T�C�Y���`�S�ɐݒ�
        If .PaperSizes(vbPRPSA4) = True Then
            .PaperSize = pprA4
        Else
            PrntDrvSW = False
            MsgBox "�p���T�C�Y��ݒ�ł��܂���ł����B", vbExclamation
            Screen.MousePointer = vbDefault   '�}�E�X������l�ɖ߂�
            Exit Sub
        End If
        '�p�����������ɐݒ�
        .Orientation = orPortrait
        If .Error <> 0 Or .Orientation = orLandscape Then
            PrntDrvSW = False
            MsgBox "�p��������ݒ�ł��܂���ł����B", vbExclamation
            Screen.MousePointer = vbDefault   '�}�E�X������l�ɖ߂�
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
'   ��ƒ��ɒ��f�����ꍇ�́A����W���u���폜���܂��B
'**********************************************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If HYOUsw = True Then
        Tyuudan = True
        If PrnMode = 1 Then VSPrinter1.KillDoc
    End If
    frmMenu.Show
End Sub

'**********************************************************************************************
'   �t�H�[���̃T�C�Y��ύX�����ꍇ�ɃR���g���[���̈ʒu��ݒ�
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
'   ���j���[�k�t�@�C���l�k�I���l
'**********************************************************************************************
Private Sub mnuEnd_Click()
    Unload Me
End Sub

'**********************************************************************************************
'   ���j���[�k�t�@�C���l�k�t�@�C���֕ۑ��l
'**********************************************************************************************
Private Sub mnuFileSave_Click()
'    Dim Datafile As String
'    Dim i As Integer, j As Integer
'    Dim SS As String, SS2 As String
'
'    CommonDialog1.FileName = "dts.csv"
'
'    CommonDialog1.CancelError = True ' CancelError �v���p�e�B��^ (True) �ɐݒ肵�܂��B
'    On Error GoTo ErrHandler
'
'    CommonDialog1.Flags = cdlOFNFileMustExist        ' Flags �v���p�e�B��ݒ肵�܂��B
'    CommonDialog1.Filter = "�b�r�u(�J���}��؂�)|*.csv|" ' ���X�g �{�b�N�X�ɕ\�������t�B���^��ݒ肵�܂��B
'    CommonDialog1.FilterIndex = 0                    ' "�e�L�X�g �t�@�C��" ������̃t�B���^�Ƃ��Ďw�肵�܂��B
'    CommonDialog1.ShowSave                           ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂��B
'    On Error GoTo 0
'
'    ' ���[�U�[���I�������t�@�C������\�����܂��B
'    Datafile = CommonDialog1.FileName
'
'    PrnMode = 2
'    ChDrive CuDir
'    ChDir CuDir
'    Open Datafile For Output As #3
'
'        SS = Trim$(kou(kou_ID, 1).TI1)
'        If kou(kou_ID, 0).no > 1 Then SS = SS & " " & Trim$(kou(kou_ID, HENI).TI2)
'        SS = SS & " �f�[�^�V�[�g"
'        Print #3, SS
'        Print #3, TNAME1 & " " & TNAME2
'
'        If Trim$(DanSet(kou_ID, 0).dan) <> "" Then
'            Print #3, "�f �� �F " & Trim$(DanSet(kou_ID, dan_ID).ti)
'        Else
'            Print #3, ""
'        End If
'
'        Print #3, "�P �� �F " & Trim$(kou(kou_ID, HENI).Yt) & " (" & Trim$(kou(kou_ID, HENI).Yu) & ")"
'        Print #3, ""
'
'        '����
'        SS = "    �v �� �� ��    "
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
'    MsgBox "�ۑ����������܂����B", vbInformation
'    Exit Sub
'ErrHandler:
'    ChDrive CuDir
'    ChDir CuDir
End Sub

'**********************************************************************************************
'   ���j���[�k�t�@�C���l�k����l
'**********************************************************************************************
Private Sub mnuPrint_Click()
    
    PrnMode = 1
    
    '���
    VSPrinter1.Action = paChoosePrintAll
End Sub

'**********************************************************************************************
'   ���j���[�k�t�@�C���l�k�������ݒ�l
'**********************************************************************************************
Private Sub mnuPrinterSet_Click()
    VSPrinter1.Action = paChoosePrinter
End Sub

'**********************************************************************************************
'   �X�N���[�� �o�[���g���āA�y�[�W���X�N���[�������A���͈͓̔��ł̌��݈ʒu���������Ƃ��ł��܂��B
'**********************************************************************************************
Private Sub scrlPage_Change()
    Dim lp As Integer
    
    scrlPage.SmallChange = VSPrinter1.PreviewPages
    scrlPage.LargeChange = scrlPage.SmallChange
    
    VSPrinter1.PreviewPage = scrlPage.Value
    
    lp = VSPrinter1.PreviewPage + VSPrinter1.PreviewPages - 1
    If lp > VSPrinter1.PageCount Then lp = VSPrinter1.PageCount
    If lp < VSPrinter1.PreviewPage Then lp = VSPrinter1.PreviewPage
    
    Label1(1).Caption = Format$(VSPrinter1.PreviewPage) & "/" & Format$(VSPrinter1.PageCount) & " �߰��"

End Sub

'**********************************************************************************************
'   ��\��ƏI����A�R���g���[�����Đݒ�
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
'   �w�b�_�[�\��
'**********************************************************************************************
Private Sub VSPrinter1_NewPage()
    Dim i As Integer, j As Integer, mco As Integer
    Dim ss As String, SS2 As String
    Dim haba1 As Single, haba2 As Single, maxcol As Integer

    '�\�̏o�͈ʒu��ݒ�

    With VSPrinter1
        .CurrentY = 1440
        .FontBold = False
        .FontItalic = False
        .FontSize = 9
        
        maxcol = DTS_Col_MAX
        haba1 = .TextWidth(String(19 + maxcol * DTS_Col_WIDTH, "-"))
        
        .FontBold = True
        .FontItalic = False
        .FontName = "�l�r ����"  'Courier
        .TextColor = vbBlack 'vbWhite
        .FontSize = 13
        
        ss = Trim$(kou(kou_ID, 1).TI1)
        ss = ss & "�v���f�[�^�V�[�g"

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

        '.Paragraph = "�P �� �F " & Trim$(kou(kou_ID, HENI).Yt) & " (" & Trim$(kou(kou_ID, HENI).Yu) & ")"
        .Paragraph = "�P �� �F " & Trim$(kou(kou_ID, HENI).Yu)
        .Paragraph = ""

        .FontBold = False  ''.FontBold = True
        .FontItalic = False
        '.FontName = "�l�r �S�V�b�N"  'Courier
        .FontSize = 9 '10
        .TextColor = vbBlack 'vbWhite
        '����

        ss = "    �v �� �� ��    "
        SS2 = "���Z�ψ�"
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


