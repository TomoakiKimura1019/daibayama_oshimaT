VERSION 5.00
Object = "{D3F92121-EFAA-4B5C-B91B-3D6A8FFD1477}#1.0#0"; "VSDraw8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeiji1 
   Caption         =   "�����v��"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19080
   Icon            =   "Keiji1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   19080
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '������
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   10290
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   27093
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   5997
            MinWidth        =   5997
            Text            =   "2013/07/10"
            TextSave        =   "2013/07/10"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���t�ݒ�"
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�P�W�O��"
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�R�O��"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�P�O��"
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   7
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�Ǘ��l�̐ݒ�ɂ���"
      Height          =   615
      Left            =   16680
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�蓮�v��"
      Height          =   615
      Left            =   16680
      TabIndex        =   4
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���ݓ����Đݒ�"
      Height          =   495
      Left            =   14880
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�蓮�v���`�F�b�N�p"
      Height          =   495
      Left            =   13800
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���j���["
      Height          =   615
      Left            =   16680
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VSDraw8LibCtl.VSDraw VSDraw1 
      Height          =   5235
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   10635
      _cx             =   18759
      _cy             =   9234
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScaleLeft       =   0
      ScaleTop        =   0
      ScaleHeight     =   1000
      ScaleWidth      =   1000
      PenColor        =   0
      PenWidth        =   0
      PenStyle        =   0
      BrushColor      =   -2147483633
      BrushStyle      =   0
      TextColor       =   -2147483640
      TextAngle       =   0
      TextAlign       =   0
      BackStyle       =   0
      LineSpacing     =   100
      EmptyColor      =   -2147483636
      PageWidth       =   0
      PageHeight      =   0
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   -1  'True
      MouseScroll     =   -1  'True
      ProportionalBars=   -1  'True
      Zoom            =   100
      ZoomMode        =   0
      KeepTextAspect  =   -1  'True
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�����ψʌv��"
      BeginProperty Font 
         Name            =   "�l�r ����"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   18975
   End
End
Attribute VB_Name = "frmKeiji1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 0
        Kkeiji.Xmax = 10
        Kkeiji.xBUN = 10
        Call WriteIni("�펞�\���ݒ�", "�w������", CStr(Kkeiji.Xmax), CurrentDir & "�v���ݒ�.ini")
        Call WriteIni("�펞�\���ݒ�", "�w��������", CStr(Kkeiji.xBUN), CurrentDir & "�v���ݒ�.ini")
        Call KeijiInit
    Case 1
        Kkeiji.Xmax = 30
        Kkeiji.xBUN = 10
        Call WriteIni("�펞�\���ݒ�", "�w������", CStr(Kkeiji.Xmax), CurrentDir & "�v���ݒ�.ini")
        Call WriteIni("�펞�\���ݒ�", "�w��������", CStr(Kkeiji.xBUN), CurrentDir & "�v���ݒ�.ini")
        Call KeijiInit
    Case 2
        Kkeiji.Xmax = 180
        Kkeiji.xBUN = 10
        Call WriteIni("�펞�\���ݒ�", "�w������", CStr(Kkeiji.Xmax), CurrentDir & "�v���ݒ�.ini")
        Call WriteIni("�펞�\���ݒ�", "�w��������", CStr(Kkeiji.xBUN), CurrentDir & "�v���ݒ�.ini")
        Call KeijiInit
    Case Else
        frmXpara.Show
    End Select
End Sub

Private Sub Command2_Click()
    Me.Visible = False
    frmMenu.Show
End Sub

Private Sub Command3_Click()
'    Dim i As Integer
'    On Error GoTo KeiERR
'
''    Form2.Height = 3225
''    Form2.ControlBox = True
''    Form2.Text1.Visible = True
''    Form2.Show
'
'    Form1.Show
'    MDY = Now
'    'Tensu% = 3
''    Tensu% = InitDT.co
''    HeikinKaisuu = InitDT.AvgCO '= 3
''    x0 = InitDT.Kx
''    y0 = InitDT.Ky
''    z0 = InitDT.Kz
''    MH = InitDT.Kmh
''    x1 = InitDT.Bx
''    y1 = InitDT.By
''    z1 = InitDT.Bz
''    For i = 1 To InitDT.co
''        H#(1, i) = PoDT(i).Hdt
''        V#(1, i) = PoDT(i).Vdt
''        S#(1, i) = PoDT(i).Sdt
''    Next
''    AZIMUTH = InitDT.HOKO
'
'
'    'GTS-8�̃X�C�b�`ON
'    Call frmGTS800A.GTS8on
'    Call frmGTS800A.GTS8init 'GTS-8�̏����ݒ�
'    '
'    Call frmGTS800A.SOKUTEI(KEISOKU.Data_path & DATA_DAT)
'
'    'GTS-8�̃X�C�b�`OFF
'    Call frmGTS800A.GTS8off
'    Exit Sub
'KeiERR:
'    Close
'    Unload Form1
'
'    If Err.Number = 10000 Then
'        MsgBox "�ʐM�G���["
'    Else
'        MsgBox "�G���[:" & Err.Number
'    End If
End Sub

Private Sub Command4_Click()
    Date = Keisoku_Time
    Time = Keisoku_Time
End Sub

'2001/05/16
Private Sub Command5_Click()
End Sub

Private Sub end_Click()
    Unload Me 'End
End Sub

Private Sub Command6_Click()
    Dim lngAPIReVal As Long
    
    'URL�����s����
    lngAPIReVal = ShellExecute(GetDesktopWindow, "open", "�D�y�h�[��������ϐ�Ɋւ���ێ��Ǘ����j.pdf", vbNullString, "", SW_SHOW)
'    lngAPIReVal = ShellExecute(GetDesktopWindow, "open", "readme.htm", vbNullString, "", SW_SHOW)
End Sub

Private Sub Form_Load()
'   Me.Height = 15360 - 420  '11000 '16590
'   Me.Width = 19200 '15000 '21210
    Left = 0 ' (Screen.Width - Me.Width) / 2
    Top = 0
'    Me.Height = 12000 '16590
'    Me.Width = 16000 '21210
'    Left = (Screen.Width - Me.Width) / 2
'    Top = 0

    Dim rc As Integer
    Dim t1 As Date, t2 As Date
    Dim i As Integer

    '2001/05/16
'    t1 = Timer
'    frmInitMsg.Show: frmInitMsg.Refresh
'    If Command$ = "TEST" Then
'        Do
'            DoEvents
'            t2 = Timer
'            If (t2 - t1) > 2 Then Exit Do
'        Loop
'    Else
'        If frmGTS800A.GTS8on = 0 Then
'            MsgBox "GTS-800�̔���������܂���B" & CStr(rc), vbCritical
'            Close
'            End
'        End If
'        Call frmGTS800A.GTS8init 'GTS-8�̏����ݒ�
'    End If
'    Unload frmInitMsg

    '�����v����ʕ\��
    'Show
    If Command$ = "CHECK" Then
        Command3.Visible = True
    Else
        Command3.Visible = False
    End If
    If Command$ = "" Then
        Command4.Visible = False
    Else
        Command4.Visible = True
    End If
    'Refresh
        
    Me.Enabled = False
    Call HyoujiInit
    keihou_L = 0
    Call DataPrint
    Me.Enabled = True
    
'    '����v�����Ԍv�Z
'    KE_intv = CDate(Lebel_intv(keihou_L) & ":00") '�C���^�[�o������
'    minTIME = DateAdd("m", 1, Now)
'    For i = 1 To 24 / Lebel_intv(keihou_L)
'        t1 = DateValue(Now) + Lebel_time(keihou_L, i)
'        If DateDiff("s", t1, Now) > 0 Then t1 = DateAdd("d", 1, t1)
'        If DateDiff("s", t1, minTIME) > 0 Then minTIME = t1
'    Next i
'    Keisoku_Time = minTIME                        '����v������
'    Call IntvWrite
'    xTextN2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'    xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
    '�o���ω��}��ʕ\��
    Kkeiji.Yti = GetIni("�펞�\���ݒ�", "�x���^�C�g��", CurrentDir & "�v���ݒ�.ini")
    Kkeiji.Yu = GetIni("�펞�\���ݒ�", "�x���P��", CurrentDir & "�v���ݒ�.ini")
    
    Kkeiji.Xmax = GetIni("�펞�\���ݒ�", "�w������", CurrentDir & "�v���ݒ�.ini")
    Kkeiji.xBUN = GetIni("�펞�\���ݒ�", "�w��������", CurrentDir & "�v���ݒ�.ini")
    If Kkeiji.Xmax = 0 Then Kkeiji.Xmax = 4: Call WriteIni("�펞�\���ݒ�", "�w������", CStr(Kkeiji.Xmax), CurrentDir & "�v���ݒ�.ini")
    If Kkeiji.xBUN = 0 Then Kkeiji.xBUN = 12: Call WriteIni("�펞�\���ݒ�", "�w��������", CStr(Kkeiji.xBUN), CurrentDir & "�v���ݒ�.ini")
    
    Kkeiji.YMIN = CSng(GetIni("�펞�\���ݒ�", "�x���ŏ��l", CurrentDir & "�v���ݒ�.ini"))
    Kkeiji.YMAX = CSng(GetIni("�펞�\���ݒ�", "�x���ő�l", CurrentDir & "�v���ݒ�.ini"))
    Kkeiji.yBUN = CSng(GetIni("�펞�\���ݒ�", "�x��������", CurrentDir & "�v���ݒ�.ini"))
    If Kkeiji.yBUN = 0 Then
        Kkeiji.yBUN = 10
        Call WriteIni("�펞�\���ݒ�", "�x��������", CStr(Kkeiji.yBUN), CurrentDir & "�v���ݒ�.ini")
    End If
    If Kkeiji.YMIN = Kkeiji.YMAX Then
        Kkeiji.YMIN = 0
        Kkeiji.YMAX = 400
        Call WriteIni("�펞�\���ݒ�", "�x���ŏ��l", CStr(Kkeiji.YMIN), CurrentDir & "�v���ݒ�.ini")
        Call WriteIni("�펞�\���ݒ�", "�x���ő�l", CStr(Kkeiji.YMAX), CurrentDir & "�v���ݒ�.ini")
    End If
    Call KeijiInit
    
'    '�����v���J�n
'    keisoku_f = False
''    Me.SetFocus
'    Timer1.Interval = 200
'    Timer1.Enabled = True
End Sub

Private Sub VSDraw1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If x < 160 Or x > 2200 Then Exit Sub
    If y < 1450 Or y > 19500 Then Exit Sub
    
    frmYPara.Show
End Sub


Private Sub HyoujiInit()
    Dim i As Integer, j As Integer
    Dim SS1 As String, SS2 As String, SS3 As String

'        xTextN2(0).Text = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
''        .Text1(1).Text = Format$(KE_intv, "           hh:nn:ss")
'        xTextN2(2).Text = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
        
        SS1 = GetIni("�펞�\���ݒ�", "�^�C�g��", CurrentDir & "�v���ݒ�.ini")
        Label1(0).Caption = SS1
        
'        For i = 0 To 3
'            SS1 = ""
'            For j = 1 To 5
'                SS2 = "���x��" & CStr(i + 1) & "���e" & CStr(j)
'                SS3 = GetIni("�펞�\���ݒ�", SS2, CurrentDir & "�v���ݒ�.ini")
'                If j > 1 Then SS3 = "  " & SS3
'
'                If j < 5 Then
'                    SS1 = SS1 & SS3 & vbCrLf
'                Else
'                    SS1 = SS1 & SS3
'                End If
'            Next j
'            xTextN1(i).Text = SS1
'            If i = 0 Then xTextN1(i).InputAreaColor = &H80FF80
'            If i = 1 Then xTextN1(i).InputAreaColor = &H80FFFF
'            If i = 2 Then xTextN1(i).InputAreaColor = &HFF80FF
'            If i = 3 Then xTextN1(i).InputAreaColor = RGB(256, 60, 60)
'        Next i
    
        '�O���t
        VSDraw1.PenWidth = 1           ' ����
        VSDraw1.FontName = "�l�r �S�V�b�N"
        VSDraw1.BrushStyle = bsTransparent

        VSDraw1.ScaleLeft = 0
        VSDraw1.ScaleWidth = 20000 '9675
        VSDraw1.ScaleTop = 10000 '9675
        VSDraw1.ScaleHeight = -10000 '-9675
End Sub


