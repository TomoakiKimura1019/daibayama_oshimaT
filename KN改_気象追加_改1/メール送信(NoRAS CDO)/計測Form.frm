VERSION 5.00
Object = "{6BD18123-F5F7-11D1-AB4F-080000047424}#1.0#0"; "RasAtl.dll"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{323DBF23-9372-4ADC-80FF-0ABA14A5F694}#4.2#0"; "xCBtnN.ocx"
Object = "{5370BF83-EEA3-4092-8CB5-7C97CE0D4324}#4.1#0"; "xLabelN.ocx"
Begin VB.Form �v��Form 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "2"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9960
   Icon            =   "�v��Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   9960
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6000
      Top             =   240
   End
   Begin VB.Frame Frame4 
      Caption         =   "�@��ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1335
      Left            =   7560
      TabIndex        =   9
      Top             =   9960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   855
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "�`�f��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "�a�f��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   870
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  '������
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   1785
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   17515
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "�v������"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �o����"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1250
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin xLabelNLib.xLabelN xLabelT 
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   825
         Width           =   2445
         _Version        =   262145
         _ExtentX        =   4313
         _ExtentY        =   503
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         Font3D          =   0
         Caption         =   "yyyy/mm/dd hh:nn:ss"
         ForeColor       =   16777215
      End
      Begin xLabelNLib.xLabelN xLabelT 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   540
         Width           =   2445
         _Version        =   262145
         _ExtentX        =   4313
         _ExtentY        =   503
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         Font3D          =   0
         Caption         =   "xLabel2"
         ForeColor       =   16777215
      End
      Begin xLabelNLib.xLabelN xLabelT 
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   2445
         _Version        =   262145
         _ExtentX        =   4313
         _ExtentY        =   529
         _StockProps     =   79
         ForeColor       =   16777215
         BackColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AlignmentH      =   1
         BorderStyle     =   4
         Font3D          =   0
         Caption         =   "xLabel2"
         ForeColor       =   16777215
      End
      Begin VB.Label Label1 
         Alignment       =   1  '�E����
         Caption         =   "�C���^�[�o��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '�E����
         Caption         =   "�O��v������"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   3
         Top             =   300
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '�E����
         Caption         =   "����v������"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   2
         Left            =   75
         TabIndex        =   2
         Top             =   900
         Width           =   1380
      End
   End
   Begin xCBtnNLib.xCmdBtnN xCmdBtn1 
      Height          =   1095
      Left            =   14400
      TabIndex        =   8
      Top             =   9480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   262146
      _ExtentX        =   2566
      _ExtentY        =   1931
      _StockProps     =   77
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionStd      =   "�x���~"
      Picture         =   "�v��Form.frx":0442
      ForeColor       =   255
      DownPicture     =   "�v��Form.frx":045E
      DisabledPicture =   "�v��Form.frx":047A
      OnPicture       =   "�v��Form.frx":0496
      PMenuCaption0   =   "���ɖ߂�(&U)"
      PEnabled0       =   -1  'True
      PHidden0        =   -1  'True
      PSeparator0     =   0   'False
      PMenuCaption1   =   "�؂���(&T)"
      PEnabled1       =   -1  'True
      PHidden1        =   -1  'True
      PSeparator1     =   0   'False
      PMenuCaption2   =   "��߰(&C)"
      PEnabled2       =   -1  'True
      PHidden2        =   -1  'True
      PSeparator2     =   0   'False
      PMenuCaption3   =   "�\��t��(&P)"
      PEnabled3       =   -1  'True
      PHidden3        =   -1  'True
      PSeparator3     =   0   'False
      PMenuCaption4   =   "�폜(&D)"
      PEnabled4       =   -1  'True
      PHidden4        =   -1  'True
      PSeparator4     =   0   'False
      PMenuCaption5   =   "���ׂđI��(&A)"
      PEnabled5       =   -1  'True
      PHidden5        =   -1  'True
      PSeparator5     =   0   'False
      PMenuCaption6   =   ""
      PEnabled6       =   -1  'True
      PHidden6        =   -1  'True
      PSeparator6     =   0   'False
      PMenuCaption7   =   ""
      PEnabled7       =   -1  'True
      PHidden7        =   -1  'True
      PSeparator7     =   0   'False
      PMenuCaption8   =   ""
      PEnabled8       =   -1  'True
      PHidden8        =   -1  'True
      PSeparator8     =   0   'False
      PMenuCaption9   =   ""
      PEnabled9       =   -1  'True
      PHidden9        =   -1  'True
      PSeparator9     =   0   'False
      PMenuCaption10  =   ""
      PEnabled10      =   -1  'True
      PHidden10       =   -1  'True
      PSeparator10    =   0   'False
      PMenuCaption11  =   ""
      PEnabled11      =   -1  'True
      PHidden11       =   -1  'True
      PSeparator11    =   0   'False
      PMenuCaption12  =   ""
      PEnabled12      =   -1  'True
      PHidden12       =   -1  'True
      PSeparator12    =   0   'False
      PMenuCaption13  =   ""
      PEnabled13      =   -1  'True
      PHidden13       =   -1  'True
      PSeparator13    =   0   'False
      PMenuCaption14  =   ""
      PEnabled14      =   -1  'True
      PHidden14       =   -1  'True
      PSeparator14    =   0   'False
      PMenuCaption15  =   ""
      PEnabled15      =   -1  'True
      PHidden15       =   -1  'True
      PSeparator15    =   0   'False
      PMenuCaption16  =   ""
      PEnabled16      =   -1  'True
      PHidden16       =   -1  'True
      PSeparator16    =   0   'False
      PMenuCaption17  =   ""
      PEnabled17      =   -1  'True
      PHidden17       =   -1  'True
      PSeparator17    =   0   'False
      PMenuCaption18  =   ""
      PEnabled18      =   -1  'True
      PHidden18       =   -1  'True
      PSeparator18    =   0   'False
      PMenuCaption19  =   ""
      PEnabled19      =   -1  'True
      PHidden19       =   -1  'True
      PSeparator19    =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "yyyy/mm/dd hh:mm:ss"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      Caption         =   "���ݓ���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
   End
   Begin RASATLCtl.RasClient RasClient1 
      Left            =   5040
      OleObjectBlob   =   "�v��Form.frx":04B2
      Top             =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   25000
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   25000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu file 
      Caption         =   "�t�@�C��"
      Begin VB.Menu end 
         Caption         =   "�I��"
      End
   End
   Begin VB.Menu mnuCrt0 
      Caption         =   "�\��"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuIntv 
      Caption         =   "�������"
   End
   Begin VB.Menu mnuSet 
      Caption         =   "�ݒ�"
      Visible         =   0   'False
      Begin VB.Menu mnuTable 
         Caption         =   "���ݒ�"
      End
   End
End
Attribute VB_Name = "�v��Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bTelDaial As Boolean  ' �_�C�A���A�b�v�J�n������True
Dim iDaialID As Integer

Dim keihouF As Boolean    '�x��t�@�C������������ True
Dim SendDataF As Boolean  '���M�t�@�C������������ True
Dim SoushinFileCK As Boolean '���M�p�t�@�C���̃`�F�b�N���ς񂾂�@True
Dim SoushinCount As Integer

Dim RASenable As Boolean  '����ڑ��� True
Dim SendComp As Boolean   '���M�I�������� True


Private Sub Command1_Click()
    'TEST�p
'            If CheckDataFile(TdsDataPath) <> 0 Then
'                iDaialID = 1
'                Call Soushin(iDaialID)
'            End If
    'keihouCK
End Sub

Private Sub end_Click()
    Unload Me 'End
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
      Unload Me
    End If
    
    Dim tp&, lt&
    On Local Error GoTo E01
    tp = CLng(GetIni("Form", "Top", mINIfile))
    If tp < 0 Then tp = 0
    lt = CLng(GetIni("Form", "left", mINIfile))
    If lt < 0 Then lt = 0
    
S01:
    Top = tp
    Left = lt
    On Local Error GoTo 0
    
    Me.Height = 2865
    Me.Width = 4545
    
    Caption = GetIni("���ꖼ", "���ꖼ", mINIfile) + " �f�[�^���M"
    
    xLabelT(1).Caption = Format$(KE_intv, "           hh:nn:ss")
    xLabelT(0).Caption = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    xLabelT(2).Caption = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    Timer1.Enabled = True
    
    Dim i As Integer
    For i = 1 To 1
        RAS(i).eName = GetIni("RAS", "eName" & i, mINIfile)
        RAS(i).User = GetIni("RAS", "User" & i, mINIfile)
        RAS(i).Pw = GetIni("RAS", "Pw" & i, mINIfile)
        Call RasInitial(i)
    Next i
    SoushinCount = 5
Exit Sub

E01:
    On Local Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rc%
    Dim i As Integer, ENDsw As Boolean, f As Integer
    Dim RetString As String
    
    On Error Resume Next
    
    If UnloadMode <= 1 Then
        If vbCancel = MsgBox("�uOK�v���N���b�N����ƁA�I�����܂��", vbOKCancel + vbExclamation, "�I���̊m�F") Then
            Cancel = True
            ENDsw = False
        Else
            ENDsw = True
        End If
    Else
        ENDsw = True
    End If
    
    If ENDsw = True Then
        
        Call WriteIni("form", "top", CStr(Top), mINIfile)
        Call WriteIni("form", "left", CStr(Left), mINIfile)
        
        '�I�����O
        f = FreeFile
        Open CurrentDIR & "PRG-event.log" For Append Lock Write As #f
            Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : �I��"
        Close #f
        
        Call IntvWrite
        Close
        
        While Forms.Count > 1
            '---�����ȊO�̃t�H�[����T���܂�
            i = 0
            While Forms(i) Is Me
                i = i + 1
            Wend
            Unload Forms(i)
        Wend
        
        '---�������g���A�����[�h���A�A�v���P�[�V�����͏I�����܂�
        Unload Me
        End
    End If
End Sub

Private Sub mnuIntv_Click()
    Dim f As Integer
    
    frmIntvNew.Show vbModal
    xLabelT(0).Caption = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    xLabelT(1).Caption = Format$(KE_intv, "           hh:nn:ss")
    xLabelT(2).Caption = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")

    '�ݒ胍�O
    f = FreeFile
    Open CurrentDIR & "PRG-event.log" For Append Lock Write As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : �C���^�[�o���ݒ�"
    Close #f
    
    '�v�����ԍĐݒ�
    Call Ktime_ck
End Sub

Private Sub RasClient1_Connected()
    Timer1.Enabled = False
    
    '�ڑ�����
    
    Dim ret%
    
    RASenable = True
    
    If keihouF = True Then
        SoushinCount = SoushinCount - 1
        If 0 < SoushinCount Then
            Call FileJoinSend
            Call SendLogMove
            Call SendPNG(ret)
            If ret = -1 Then
                Call DelFile(PNGpath & "*.png")
            End If
        Else
            SoushinCount = 0
        End If
        keihouF = False
    Else
        SoushinCount = 5
    End If
    
    If SendDataF = True Then
        Call FindDataFile(2, LZDataPath)
        Call SendPNG(ret)
        SendDataF = False
    End If
    
    SendComp = True
    
    RasClient1.HangUp  '�d�b��؂�
    Call Sleep(2000)
    StatusBar1.Panels(1).Text = ""
    SoushinFileCK = False
    Timer1.Enabled = True
Exit Sub
            'Shape1.FillColor = RGB(0, &HFF, 0)
'            StatusBar1.Panels(1).Text = "�ڑ���"
            'Call SendFTP(ret)
'            If RasClient1.Active = True Then
'                RasClient1.HangUp  '�d�b��؂�
'                Call Sleep(2000)
'                ConnectCK = 0
'                StatusBar1.Panels(1).Text = ""
'            End If
'            If ret = -1 Then
'                'Call DelFile("\FTP\" & Trim(henkanTBL(1).FileName) & ".dat")
'                Call AllFileDelete
'            End If
'            Z_Keisoku_Time = Keisoku_Time
'            Keisoku_Time = CDate(Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")) + KE_intv
'            If rSettei = True Then
'                Keisoku_Time = CDate(GetIni("�v������", "����v������", MyAppPath & "settei.ini"))
'                KE_intv = CDate(GetIni("�v������", "�v���C���^�[�o��", MyAppPath & "settei.ini"))
'                Call IntvWrite
'                Text1(1) = Format$(KE_intv, "           hh:nn:ss")
'                Text1(2) = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'                Call Ktime_ck: Text1(2) = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
'                KEISOKU.Waittime = CInt(GetIni("�v��", "�҂�", MyAppPath & "settei.ini"))
'                Call WriteIni("�v��", "�҂�", CStr(KEISOKU.Waittime), MyAppPath & "fSoushin.ini")
'                Call DelFile(App.Path & "\settei.ini")
'                rSettei = False
'            End If
'    Call FTPrw
End Sub

Private Sub RasClient1_StatusChange(StatusMsg As String, ByVal StatusCode As Long)
    StatusBar1.Panels(1).Text = str$(StatusCode) & " " & StatusMsg
End Sub

Private Sub RasClient1_StatusError(ErrorMsg As String, ByVal ErrorCode As Long)
    'Form2.Label1 = Str$(ErrorCode) & ErrorMsg
    'Form2.CmdCancel.Caption = "OK"
    Dim f As Integer
    
    ConnectCK = -3
    If ErrorCode = 676 Then ConnectCK = -2    '�޼ް
    If ErrorCode = 678 Then ConnectCK = -6    '����������܂���
    If ErrorCode = 633 Then ConnectCK = -7    '�߰Ă͊��Ɏg�p�����A�Ӱ� ���� �޲�ٱ�Ăɑ΂��č\������Ă��܂���B
    
    StatusBar1.Panels(1).Text = str$(ErrorCode) & " " & ErrorMsg
    
    f = FreeFile
    Open CurrentDIR & "modem-err.dat" For Append As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " :";
        Print #f, str$(ErrorCode) & " " & ErrorMsg
    Close #f
End Sub

Private Function RasStart(ByVal id As Integer) As Boolean
'    Dim Ent As Entry
'    Set Ent = RasClient1.Items.Item(iEntNo)
'    tx5 = Ent.UserName
'    tx6 = Ent.Password
'    Set Ent = Nothing
    
    'NT�̏ꍇRasmon.exe���N�����܂��B
    On Error Resume Next
    If m_Rasmon = 0 Then
        m_Rasmon = Shell("rasmon.exe", vbNormalFocus)
    End If
    On Error GoTo 0
    
    With �v��Form.RasClient1
        .ItemIndex = RAS(id).iEntNo
        '�v���p�e�B�[���Z�b�g���܂��B
        .ReDialTimes = 1
        .ReDialInterval = 10
        ''Form2.Caption = RasClient1.EntryName
        .UserName = "a545352322@p.auone-net.jp" 'RAS(id).User '"���[�U�[����ݒ�"
        .Password = "yo2803ks" 'RAS(id).Pw '"�p�X���[�h��ݒ�"
    ''    If Check1.Value = 0 Then
    ''        .UseSpValue = True
    ''        .SpTelephoneNumber = Text2
    ''        .SpDomainName = Text3
    ''        .SpCallBackNumber = Text4
    ''    Else
            '�_�C�A���A�b�v�l�b�g���[�N�ɂ��炩���߃Z�b�g���ꂽ�l���g�p�����
            .UseSpValue = False
    ''    End If
        
        '�ڑ����܂��B
        If .Connect(Me.hWnd) = -1 Then
            '�ڑ������J�n
            RasStart = True
        Else
            ''�ڑ������Ɏ��s
            RasStart = False
        End If
    End With
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    Dim Tstime As String
    Dim keisoku_f As Boolean
    Dim ret As Integer
    
    Tstime = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
    
    If SoushinFileCK = False Then Call tmSub
    If RasClient1.Active = True Then
        RasClient1.HangUp  '�d�b��؂�
        Call Sleep(3000)
    End If
'    If RASenable = True And SendComp = True Then
'        RasClient1.HangUp  '�d�b��؂�
'        Call Sleep(2000)
'        RASenable = False
'        SendComp = False
'    End If
    
    '�v�����Ԃ��������ꍇ
    If DateDiff("s", Keisoku_Time, Now) > 0 Then
        Keisoku_Time = T_ajt(Z_Keisoku_Time, KE_intv)
        xLabelT(0).Caption = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
        xLabelT(2).Caption = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
    End If

    Timer1.Enabled = True
End Sub

Private Sub tmSub()
    
    '�x��f�[�^�t�@�C������������A���̃f�[�^�𑗐M
    keihouF = keihouCK
    
    '�v���f�[�^�t�H���_�̒��g����������A���̃f�[�^�𑗐M
    If CheckDataFile(LZDataPath) <> 0 Then
        SendDataF = True
    End If

    If keihouF = True Or SendDataF = True Then
        '�d�b��������
        If RasStart(2) = False Then Exit Sub
        SoushinFileCK = True
    End If

    On Local Error Resume Next
    Dim f As Integer
    If keihouF = True Then
        f = FreeFile
        Open BkeiPath & "keihou.txt" For Output As #f
        Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss")
        Close #f
    End If
    On Local Error GoTo 0
Exit Sub

    
    Dim Tstime As String
    Dim keisoku_f As Boolean
    
        Tstime = Format$(Now, "yyyy/mm/dd hh:nn:ss")
        Label3 = Format$(Now, "yyyy/mm/dd hh:mm:ss")
        
        If Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss") = Tstime Then keisoku_f = True
        
        If keisoku_f = True Then
            StatusBar1.Panels(1).Text = "*** �ʐM�� ***"
            
            '�x��f�[�^�t�@�C������������A���̃f�[�^�𑗐M
            keihouF = keihouCK
            
            '�v���f�[�^�t�H���_�̒��g����������A���̃f�[�^�𑗐M
            If CheckDataFile(LZDataPath) <> 0 Then
                SendDataF = True
'                iDaialID = 1
'                Call Soushin(iDaialID)
            End If

            SoushinFileCK = True

            If keihouF = True Or SendDataF = True Then
                If RasStart(2) = False Then strData = "": �v��Form.StatusBar1.Panels(1).Text = "": Exit Sub
            End If

            
            If keisoku_f = True Then
                keisoku_f = False
                Z_Keisoku_Time = Keisoku_Time
                Keisoku_Time = CDate(Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")) + KE_intv
            End If
            
            Call IntvWrite
            
            Enabled = True
            
            xLabelT(0).Caption = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            xLabelT(2).Caption = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            
        End If

'        If RasClient1.Active = True Then
'            RasClient1.HangUp
'            Sleep (2000)
'        End If
        
        '�v�����Ԃ��������ꍇ
        If DateDiff("s", Keisoku_Time, Now) > 0 Then
            Keisoku_Time = T_ajt(Z_Keisoku_Time, KE_intv)
            xLabelT(0).Caption = Format$(Z_Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
            xLabelT(2).Caption = Format$(Keisoku_Time, "yyyy/mm/dd hh:nn:ss")
        End If
'''        '2001/5/14
'''        If DateDiff("s", Kiroku_Time, Now) > 0 Then
'''            Kiroku_Time = T_ajt(Z_Kiroku_Time, KI_intv)
'''            .Text1(3).Text = Format(Z_Kiroku_Time, "yyyy/mm/dd hh:nn:ss")
'''            .Text1(5).Text = Format(Kiroku_Time, "yyyy/mm/dd hh:nn:ss")
'''        End If

End Sub

Private Sub Soushin(ByVal id As Integer)
    Dim ModemSW As Boolean
    Dim t1 As Date, t2 As Date
    
    ConnectCK = 0
    ModemSW = RasStart(id) '�d�b����
    If ModemSW = False Then strData = "": �v��Form.StatusBar1.Panels(1).Text = "": Exit Sub
    t1 = Now
    Do
        DoEvents
        If ConnectCK = 1 Then Exit Do            '��M����
        If ConnectCK = 2 Then Exit Do            '�r�W�[
        If ConnectCK = 3 Then Exit Do            '����G���[
'        If ConnectCK = 4 Then Exit Do            '��M���s
'        If ConnectCK = 5 Then Exit Do            '�x��ʒm
        If ConnectCK = 6 Then Exit Do            '����������܂���B
        
        t2 = Now
        
        '�T���҂��Ă��C�x���g���N���Ȃ���Γd�b��؂�B
        If DateDiff("s", DateAdd("s", 600, t1), t2) > 0 Then Exit Do
    Loop
    
        StatusBar1.Panels(1).Text = "����ؒf�J�n..."
        RasClient1.HangUp  '�d�b��؂�
        StatusBar1.Panels(1).Text = "����ؒf..."
        Call Sleep(2000)
        ConnectCK = 0
        StatusBar1.Panels(1).Text = ""
End Sub

Private Function keihouCK() As Boolean
'�x��e�L�X�g�t�@�C�������݂��邩
    keihouCK = False
    Dim kPath As String
    kPath = KEISOKU.keihou_path & "Send000.txt"
    If FileExists(kPath) = True Then
        keihouCK = True
'        iDaialID = 2
'        Call Soushin(iDaialID)
        'Call FileJoinSend
        'Call SendLogMove
    End If
End Function

Private Sub FileJoinSend()
'////////////////////////////////////////////////////////////////
'�t�@�C����A�����āA���M���b�Z�[�W�e�L�X�g���쐬 & ���M

    Dim i As Integer, f As Integer
    Dim f2 As Integer
    Dim f3 As Integer
    Dim L As String
    Dim SS1$
    
    Dim ntm As Date
    
    strData = "" '"�ȉ��̃f�[�^���Ǘ��l���I�[�o�[���Ă��܂��B" & vbCrLf
    f3 = FreeFile
    Open KEISOKU.keihou_path & "Send000.txt" For Input As #f3
    Do While Not (EOF(f3))
        Input #f3, L
        If Left$(L, 1) <> ";" Then
            strData = strData & L & vbCrLf
        End If
    Loop
    Close #f3
    
'    If MailTabl.JyusinSW = 1 Then
'        Call MailRead(MailTabl) '���[����M�����s
'    End If
    
'Exit Sub
    
    Call MailSend(MailTabl) '���[�����M�����s
    ntm = Now
    Call WriteIni("���[�����M", "�ŏI���M����", Format$(ntm, "yyyy/mm/dd hh:nn:ss"), mINIfile)
'    Label2(2) = Format$(ntm, "yyyy/mm/dd hh:nn:ss")
    
    strData = ""
    
    StatusBar1.Panels(1).Text = ""
    
End Sub

Private Sub SendLogMove()
    Dim f As Integer
    Dim f2 As Integer
    Dim sc As Integer
    Dim bf As String
    
    sc = GetIni("���[�����M", "sc", mINIfile)
    sc = sc + 1
    f = FreeFile
    Open KEISOKU.keihou_path & "send000.txt" For Input As #f
    f2 = FreeFile
    Open App.Path & "\sendlog\send" & Format(sc, "0000") & ".txt" For Output As #f2
    Do While Not (EOF(f))
        Line Input #f, bf
        Print #f2, bf
    Loop
    Close #f2
    Close #f
    Call WriteIni("���[�����M", "sc", (sc), mINIfile)

    Call DelFile(KEISOKU.keihou_path & "send000.txt")
End Sub


Private Sub RasInitial(ByVal id As Integer)
    'RAS�G���g���[����ړI�̐ڑ����T��
    Dim i As Integer
    
    With �v��Form
    If .RasClient1.ItemCount = 0 Then
        '�G���g���[����
        MsgBox "�G���g���[����": End 'Stop
    Else
            For i = 0 To .RasClient1.ItemCount - 1
                If .RasClient1.Item(i) = RAS(id).eName Then
                    '�G���g���[�ԍ���ێ�
                    RAS(id).iEntNo = i
                    Exit For
    '            Else
    '                '�ڑ��悪������Ȃ�
    '                MsgBox "'�ڑ��悪������Ȃ�": End 'Stop
                End If
            Next i
            'If i = RasClient1.ItemCount Then
                'MsgBox "'�ڑ��悪������Ȃ�": End
            'End If
    End If
    End With
End Sub


