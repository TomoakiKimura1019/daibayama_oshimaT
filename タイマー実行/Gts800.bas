Attribute VB_Name = "GTS800A"
Option Explicit

Public RSMerr As Integer
Public Const ACK As Byte = 6
Public Const EXT As Byte = 3

'Public H(2, 16) As Double
'Public V(2, 16) As Double
'Public S(2, 16) As Double

Public DH(17) As Double
Public DV(17) As Double
Public XD(20) As Double '������W
Public YD(20) As Double
Public ZD(20) As Double
Public xo(16) As Double '������W�i�O�l�j
Public yo(16) As Double
Public zo(16) As Double

Public XN As Double, YN As Double, ZN As Double         '������W�i�V�l�j
Public dx As Double, dy As Double, dz As Double         '�ړ���

'Public x0 As Double, y0 As Double, z0 As Double, MH#    '��B�_���W�A��B��
'Public x1 As Double, y1 As Double, z1 As Double         '�㎋�_���W
'Public HeikinKaisuu As Integer
'Public Tensu As Integer
'Public AZIMUTH#

Public Const RAD As Double = 3.14159265358979 / 180#
Public iCount As Long

Public ssCmd As String
Public srCmd As String
Public MDY As Date      '�v������

Sub Fin()
    Dim rc As Integer, i As Integer
    
    Close
    
    Unload RsctlFrm
    End
End Sub

Sub KijyunIn()
'�ϑ��_�̏����ݒ�
'    Dim AZIMUTH#
    Dim h1 As Double, v1 As Double, s1 As Double, c As Integer, rl As String
    Dim i As Integer

'   ��f�[�^���i�[����t�@�C���� =
    Open "Text1.Text" For Output As #2
''''   �@�B�_  X���W  X0
''''           Y���W  Y0
''''           Z���W  Z0
''''           ��B�� MH
'''    x0 = Val(Text2(0).Text)
'''    y0 = Val(Text2(1).Text)
'''    z0 = Val(Text2(2).Text)
'''    MH = Val(Text6.Text)
'''    Print #2, x0, y0, z0, MH
''''   �㎋�_ X���W  X1
''''          Y���W  Y1
''''          Z���W  Z1
'''    x1 = Val(Text3(0).Text)
'''    y1 = Val(Text3(1).Text)
'''    z1 = Val(Text3(2).Text)
'''    Print #2, x1, y1, z1
'''    Call CalcAzimuth
'''    'LOCATE 15, 10: Print " �����p : "; AZIMUTH; "(deg)"
'''    Text6.Text = AZIMUTH
    
    MsgBox "BACK(1����)�𐳂Ŏ������AENTER�������Ă��������B", vbOKOnly, ""

'���_�Z�b�g
    Call SendCmd("ZB1" + "+0000000d")   '�����p�̐ݒ� 0000
    Call DataIn(h1, v1, s1, c, rl)
    PoDT.H(1, 1) = h1
    PoDT.V(1, 1) = v1
    PoDT.s(1, 1) = s1
    RsctlFrm.StatusBar1.Panels(1).Text = "H:" & Format(h1#, "000000000") & " V:" & Format(v1#, "000000000") & " S:" & Format(s1#, "0000000")
''    Text5(0).Text = H1#
''    Text5(1).Text = V1#
''    Text5(2).Text = S1#
    Print #2, h1, v1, s1
    
'�w��̓_���������ʒu�̓ǂݍ���
    For i = 1 To InitDT.Tensu - 1
        MsgBox i + 1 & "�����𐳂Ŏ������AENTER�������Ă��������B", vbOKOnly, ""
        Call DataIn(h1, v1, s1, c, rl)
        PoDT.H(1, i + 1) = h1
        PoDT.V(1, i + 1) = v1
        PoDT.s(1, i + 1) = s1
        RsctlFrm.StatusBar1.Panels(1).Text = "H:" & Format(h1, "000000000") & " V:" & Format(v1, "000000000") & " S:" & Format(s1, "0000000")
'        Text5(0).Text = H1#
'        Text5(1).Text = V1#
'        Text5(2).Text = S1#
        Print #2, h1, v1, s1
    Next i
    Close #2
End Sub

Sub SOKUTEI(FileName As String)
'TUIKAI_KAN(FileName As String)
'�v�����s
    Dim h1 As Double, v1 As Double, s1 As Double, c As Integer, rl As String
    Dim byo As Double, dms As String, HH1 As String, VV1 As String
    Dim A As String
    Dim j As Integer, i As Integer
    Dim SW(10) As Boolean, poNO As Integer, TryCo As Integer, TryMAX As Integer
    Dim ckPO_j As Double, ckPO_i As Double
    Dim Rec As Integer, f As Integer
    
'   �f�[�^���i�[����t�@�C����
    Open "test_d.dat" For Output As #3
    '���_�Z�b�g
                RsctlFrm.StatusBar1.Panels(1).Text = "���_�Z�b�g"
    Call TiltIn(h1#, v1#)      '�p�x�v��
    byo# = PoDT.H(1, 1) - h1#
    If byo# > 180# * 3600# Then byo# = byo# - 360# * 3600#
    If byo# < -180# * 3600# Then byo# = byo# + 360# * 3600#
    Call BYOtoDMS(byo#, dms$)
    HH1$ = Left$(dms$, 8)
    byo# = v1# - PoDT.V(1, 1)
    Call BYOtoDMS(byo#, dms$)
    VV1$ = Left$(dms$, 8)
                RsctlFrm.StatusBar1.Panels(1).Text = "���_�֐���"
    
    A$ = "T13" + VV1$ + HH1$ + "d"  '���_�֐���ǔ�
    Call SendCmd(A$)
                RsctlFrm.StatusBar1.Panels(1).Text = "���_�ǔ�"
    Call senkaiWAIT

''    A$ = "T10" + VV1$ + HH1$ + "d"  '���_�֐���
''    Call SendCmd(A$)
''                Form2.StatusBar1.Panels(1).Text = "���_�ǔ�"
''    Call senkaiWAIT
''    Call SendCmd("T34")             '�ǔ�
''    Call senkaiWAIT
                
                RsctlFrm.StatusBar1.Panels(1).Text = "�ǔ�OK"
    Call SendCmd("T30")     '�X�^���o�C
    
'''    Call SendCmd("ZB1" + "+0000000d")   '�����p�̐ݒ�
    
    j% = 0
    Print #3, Chr$(34); "H"; Chr$(34); ","; Chr$(34); "M"; Chr$(34); ","; Chr$(34); "S"; Chr$(34); ",";
    Print #3, Chr$(34); ""; Chr$(34); ",";
    Print #3, Chr$(34); ""; Chr$(34); ",";
    Print #3, Chr$(34); "X "; Chr$(34); ","; Chr$(34); "Y"; Chr$(34); ","; Chr$(34); " Z "; Chr$(34); ",";
    Print #3, Chr$(34); " SD "; Chr$(34); ",";
    Print #3, Chr$(34); "HA d"; Chr$(34); ","; Chr$(34); "m"; Chr$(34); ","; Chr$(34); "s"; Chr$(34); ",";
    Print #3, Chr$(34); "VA d"; Chr$(34); ","; Chr$(34); "m"; Chr$(34); ","; Chr$(34); "s"; Chr$(34)
    
                RsctlFrm.StatusBar1.Panels(1).Text = "���_�v���J�n"
'    While (1)
        'DoEvents
        '''iCount = iCount + 1
        Call SendCmd("T34")     '�ǔ��R�}���h
        Rec = DataInCN(j% + 1)
        
        If Rec = 0 Then
            '�������ǂ߂Ȃ��������́A�t�H�[�����J��
            Call SendCmd("T30")     '�X�^���o�C
            '���O
            If frmCLOSE.MSGfrm = False Then frmMSG.Show
            f = FreeFile
            Open CurrentDIR & "PRG-event.log" For Append Lock Write As #f
                Print #f, Format$(Now, "yyyy/mm/dd hh:nn:ss") & " : ���_�m�F���o���܂���ł����B"
            Close #f
            
            Close #3
            Exit Sub
        Else
            '�������ǂ߂����́A�t�H�[�������i���Ƀt�H�[�����J���Ă��鎞�̂݁j
            If frmCLOSE.MSGfrm = True Then Unload frmMSG
        End If
            
            Call DataIn(h1#, v1#, s1#, c%, rl$)
            PoDT.H(2, 1) = h1#
            PoDT.V(2, 1) = v1#
            PoDT.s(2, 1) = s1#
        
        Call SendCmd("T30")     '�X�^���o�C
        Call SendCmd("T200")     '�Ή񓮍�R�}���h
        Call senkaiWAIT
        Call SendCmd("T34")     '�ǔ��R�}���h
        Call DataInCN(j% + 1)
        Call SendCmd("T30")     '�X�^���o�C
        Call DataCal(j% + 1)
        Call senkaiWAIT
        
        
'            Close #3
'            Exit Sub
        For j% = 1 To InitDT.Tensu - 1
            SW(j + 1) = False
        Next j%
        
        TryMAX = 5 '�v�����[�v��
        For TryCo = 1 To TryMAX
            For j% = 1 To InitDT.Tensu - 1
                If SW(j + 1) = True Then GoTo skip_1
                    If j = 1 Then
                        RsctlFrm.StatusBar1.Panels(1).Text = "��_�v���J�n(��)"
                    Else
                        RsctlFrm.StatusBar1.Panels(1).Text = j - 1 & " �_�ڌv���J�n(��)"
                    End If
                Call DataIn(h1#, v1#, s1#, c%, rl$)
                byo# = PoDT.H(1, j% + 1) - h1#
                If byo# > 180# * 3600# Then byo# = byo# - 360# * 3600#
                If byo# < -180# * 3600# Then byo# = byo# + 360# * 3600#
                Call BYOtoDMS(byo#, dms$)
                HH1$ = Left$(dms$, 8)
                byo# = v1# - PoDT.V(1, j% + 1)
                Call BYOtoDMS(byo#, dms$)
                VV1$ = Left$(dms$, 8)
                A$ = "T13" + VV1$ + HH1$ + "d"  '����&�T�[�`�R�}���h
                Call SendCmd(A$)
                If senkaiWAIT = 0 Then
                    Err.Raise 10000 '�G���[����
                    Exit Sub
                End If
                ''call senkaiWAIT
    
                DataInCN (j% + 1)
                    Call DataIn(h1#, v1#, s1#, c%, rl$)
                
                If PoDT.s(2, j% + 1) = 0 Then ckPO_j = PoDT.s(1, j% + 1) Else ckPO_j = PoDT.s(2, j% + 1)
                poNO = 0
                If Abs(ckPO_j - s1#) > 1# Then
                    '�O��̋�����10m�ȏ㗣��Ă�����T���āA10m�ȓ��Ɍ���������
                    '���̔z��ɓ����
                    For i = 1 To InitDT.Tensu - 1
                        If PoDT.s(2, i% + 1) = 0 Then ckPO_i = PoDT.s(1, i% + 1) Else ckPO_i = PoDT.s(2, i% + 1)
                        If Abs(ckPO_i - s1#) < 1# And SW(i + 1) = False Then
                            poNO = i + 1
                            SW(i + 1) = True
                            Exit For
                        End If
                    Next i
                Else
                    SW(j + 1) = True
                    poNO = j + 1
                End If
If poNO = 0 Then GoTo skip_1
                PoDT.H(2, poNO) = h1#
                PoDT.V(2, poNO) = v1#
                PoDT.s(2, poNO) = s1#
                
                Call SendCmd("T30")             '�X�^���o�C
                Call SendCmd("T200")            '�Ή񓮍�R�}���h
                    If poNO = 2 Then
                        RsctlFrm.StatusBar1.Panels(1).Text = "��_�v���J�n(��)"
                    Else
                        RsctlFrm.StatusBar1.Panels(1).Text = CStr(poNO - 2) & " �_�ڌv���J�n(��)"
                    End If
                Call senkaiWAIT
                Call SendCmd("T34")             '�ǔ��R�}���h
                DataInCN (poNO)
                Call SendCmd("T30")             '�X�^���o�C
                
                Call DataCal(poNO)
                Call senkaiWAIT
    
    '            PoDT.H(2, j% + 1) = H1#
    '            PoDT.V(2, j% + 1) = V1#
    '            PoDT.S(2, j% + 1) = S1#
    '
    '
    '            Call SendCmd("T30")             '�X�^���o�C
    '            Call SendCmd("T200")            '�Ή񓮍�R�}���h
    '                Form2.StatusBar1.Panels(1).Text = j + 1 & " �_�ڌv���J�n(��)"
    '            Call senkaiWAIT
    '            Call SendCmd("T34")             '�ǔ��R�}���h
    '            DataInCN (j% + 1)
    '            Call SendCmd("T30")             '�X�^���o�C
    '
    '            Call DataCal(j% + 1)
    '            Call senkaiWAIT
skip_1:
            Next j%
        Next TryCo
        
                RsctlFrm.StatusBar1.Panels(1).Text = "�v���I��"
        j% = 0
        Call DataIn(h1#, v1#, s1#, c%, rl$)
        byo# = PoDT.H(1, 1) - h1#
        If byo# > 180# * 3600# Then byo# = byo# - 360# * 3600#
        If byo# < -180# * 3600# Then byo# = byo# + 360# * 3600#
        Call BYOtoDMS(byo#, dms$)
        HH1$ = Left$(dms$, 8)
        byo# = v1# - PoDT.V(1, 1)
        Call BYOtoDMS(byo#, dms$)
        VV1$ = Left$(dms$, 8)
                RsctlFrm.StatusBar1.Panels(1).Text = "���_�֕��A"
        'A$ = "T10" + VV1$ + HH1$ + "d"      '�w��p�x�Ő���
        A$ = "T13" + VV1$ + HH1$ + "d"  '���_�֐���ǔ�
        Call SendCmd(A$)
        Call senkaiWAIT
        Call SendCmd("T30")     '�X�^���o�C
'    Wend
    Close #3
                RsctlFrm.StatusBar1.Panels(1).Text = "�f�[�^�ۑ�"
    Call ZahyoWrite(FileName)
    Call HouiWrite
                RsctlFrm.StatusBar1.Panels(1).Text = "�ۑ��I��"
End Sub

Function ComLinput(rs As String) As Integer
    'COM�|�[�g����LF������܂œǂݍ���
    Dim dummy As Integer
    Dim rc As String
    Dim iv As Date
    
    srCmd = ""
    RSMerr = 0
    ComLinput = 1
    rs$ = ""
    iv = Now
    Do
        DoEvents
        If DateDiff("s", iv, Now) > 7 Then
            Exit Do
        End If
        rc$ = RsctlFrm.VBMCom1.RecvString(1)
        If RSMerr <> 0 Then
            rs$ = ""
            Exit Function
        End If
        If rc$ = Chr$(&HD) Then rc$ = ""
        If Right$(rc$, 1) = Chr$(&HA) Then rc$ = "": ComLinput = 0: Exit Do
        rs$ = rs$ + rc$
'        Debug.Print rs$
    Loop
'        Debug.Print rs$
        If Command$ <> "" Then Form1.Text1.SelText = rs$ & vbCrLf
        'srCmd = rs$: Form2.StatusBar1.Panels(1).Text = srCmd
End Function

Private Sub BCCcal(A As String, BC As String)
    'BCC�v�Z
    Dim BCC As Integer, i As Integer
    
    BCC = 0
    For i = 1 To Len(A)
        BCC = BCC Xor Asc(Mid$(A, i, 1))
    Next i
    BC = Right$("000" & Right$(str$(BCC), Len(str$(BCC)) - 1), 3)
End Sub

Sub GTS8init()
    On Error GoTo InitErr
    
    'GTS-8�̏����ݒ�
    If SendCmd("ZB23") = 0 Then GoTo InitErr        'EDM���[�h�ݒ�(�t�@�C��)
    If SendCmd("ZB4+") = 0 Then GoTo InitErr        '�����p�����ݒ�(�E�܂��=�v���X)
    If SendCmd("ZB52") = 0 Then GoTo InitErr        '�`���g�ݒ�(����&������␳)
    If SendCmd("ZB61") = 0 Then GoTo InitErr        '�V��0�Z�b�g
    
'    If SendCmd("ZC2010010d") = 0 Then GoTo InitErr  '�T�[�`�ݒ�?
'    If SendCmd("ZD30010") = 0 Then GoTo InitErr     '�E�F�C�g���Ԑݒ�(10�b)
    If SendCmd(InitDT.Serch) = 0 Then GoTo InitErr  '�T�[�`�ݒ�?
    If SendCmd(InitDT.Wait) = 0 Then GoTo InitErr     '�E�F�C�g���Ԑݒ�(10�b)
    
    If SendCmd("ZD50") = 0 Then GoTo InitErr        '�g���b�L���O�C���W�P�[�^�ݒ�(�I�t)
    On Error GoTo 0
    Exit Sub

InitErr:
    MsgBox "�ݒ�R�}���h����M����Ȃ��B" & vbCrLf & "�@��𒲍����Ă��������B", vbCritical, "�ʐM��Q"
    Call Fin
End Sub

Function SendCmd(cmd As String) As Integer
'GTS-8�Ƀf�[�^���M
' 0:NG
'-1:OK
    Dim Srbuf As String
    Dim rc As Integer
    Dim RT As String, BC As String
    Dim ic As Date
    
    SendCmd = 0
    ssCmd = ""
    'BCC�v�Z
    Call BCCcal(cmd, BC)
    ic = Now
    Do
        If DateDiff("s", ic, Now) > 20 Then
            SendCmd = 0
            Err.Raise 10000 '�G���[����
            Exit Do
        End If
        Srbuf = cmd & BC & Chr(EXT) & vbCrLf
        '���M
        rc = RsctlFrm.VBMCom1.SendString(Srbuf)
            'Debug.Print Srbuf; " : ";
            If Command$ <> "" Then Form1.Text1.SelText = Srbuf & " : "
        'ACK��M
        rc = ComLinput(RT$)
    
        If RT$ = Chr(ACK) & "006" & Chr(EXT) Then
            SendCmd = -1
            Exit Do
        End If
    Loop
    
End Function

Sub CalcAzimuth()
   '�����p�̌v�Z
    Dim ax As Double, ay As Double
'    ax = InitDT.x1 - InitDT.x0
'    ay = InitDT.y1 - InitDT.y0
    ax = XD(2) - XD(1)
    ay = YD(2) - YD(1)
    
    If (ax = 0# And ay = 0#) Then
        InitDT.AZIMUTH = 0#
    ElseIf (ax = 0#) Then
        If (ay > 0#) Then
            InitDT.AZIMUTH = 90#
        Else
            InitDT.AZIMUTH = 270#
        End If
    Else
        InitDT.AZIMUTH = Atn(ay / ax) / RAD#
    End If
    If (ax < 0#) Then InitDT.AZIMUTH = 180# + InitDT.AZIMUTH
    If (ax > 0# And ay < 0#) Then InitDT.AZIMUTH = 360# + InitDT.AZIMUTH

    InitDT.AZIMUTH = -InitDT.AZIMUTH
End Sub

Function DataIn(h1 As Double, v1 As Double, s1 As Double, c As Integer, rl As String) As Integer
'�v�����s�i�f�[�^�v���j
' 0:NG
'-1:OK
    Dim q As Integer, sData As String, rc As Integer
    Dim h2 As String, v2 As String, S2 As String
    Dim dms As String, byo As Double
    
'0        1         2         3         4         5         6         7
'1234567890123456789012345678901234567890123456789012345678901234567890
'Q+011784812m08520300+12030400d+011745724t15+0000+025000r121
'|           |<  8 >|V                                 ^
' |<  10  >|SD       |<  9  >|H                         |R or L
    
    h1 = 0#
    v1 = 0#
    s1 = 0#
    c = 4
    rl = ""
    
    DataIn = 0
    Call SendCmd("C11")       '�΋������[�h�̃f�[�^�v��
    rc = ComLinput(sData)
    q = InStr(sData, "Q")
    If q = 0 Then
    '   �f�[�^�擾�s�\
        Exit Function
    End If
    If Len(sData) < 59 Then
    '   �f�[�^�擾�s�\
        Exit Function
    End If
    
    h2 = Mid$(sData, q% + 20, 9)
    v2 = Mid$(sData, q% + 12, 8)
    S2 = Mid$(sData, q% + 1, 10)
    c = Val(Mid$(sData, q% + 54, 1))
    rl = Mid$(sData, q% + 55, 1)
    dms = h2
    Call DMStoBYO(byo, dms)
    h1 = byo
    dms = " " & v2
    Call DMStoBYO(byo, dms)
    v1 = byo
    s1 = Val(S2) / 10000#
    
    DataIn = -1
End Function

Sub DMStoBYO(byo As Double, dms As String)
   '�x���b��b�ɕϊ�
    Dim sg As Integer, DD As Integer, fu As Integer, bb As Integer
    sg = Sgn(Val(dms))
    DD = Val(Mid$(dms, 2, 3))
    fu = Val(Mid$(dms, 5, 2))
    bb = Val(Mid$(dms, 7, 3))
    byo = sg * (DD * 3600# + fu * 60# + bb / 10#)
End Sub

Function DataInCN(Pnum As Integer) As Integer
'�w��̉񐔃f�[�^������s��
' 0:NG
'-1:OK
    Dim h1 As Double, v1 As Double, s1 As Double, c As Integer, rl As String
    Dim i As Integer, q As Long
    Dim tm As String, st As Integer, ccn As Integer

    DataInCN = 0
    ccn = 0
    s1 = 0#
    While (s1 = 0)
        DoEvents
        ccn = ccn + 1: If ccn > 20 Then Exit Function
        Call DataIn(h1, v1, s1, c, rl)
    Wend
    For i% = 1 To 2 '6
        Call DataIn(h1, v1, s1, c, rl)
        Call SecWait(1)
        'For Q& = 1 To 45000: Next
    Next i%
    
    '�L�^
    i = 1
    Do
        DoEvents
    ''''For i% = 1 To HeikinKaisuu
        Call DataIn(h1, v1, s1, c, rl)
        If (s1 <> 0#) Then
            Call XyzCal(i, h1, v1, s1, rl)
            tm = Time$
            Call DataWRITE(tm, Pnum, rl, XD(i), YD(i), ZD(i), h1, v1, s1)
        Else
            i = i - 1
        End If
        Call SecWait(1)
'''        For Q& = 1 To 45000: Next
    ''''Next i%
        i = i + 1
        If i > InitDT.HeikinKaisuu Then Exit Do
    Loop
    DataInCN = -1
End Function

Sub DataWRITE(tm As String, Pnum As Integer, rl As String, _
              XF As Double, YF As Double, ZF As Double, h1 As Double, v1 As Double, s1 As Double)

'�S�Ă̊ϑ���̃f�[�^���L�^����B
    Dim hdo As Integer, hfun As Integer, hbyo As Integer, vdo As Integer, vfun As Integer, vbyo As Integer
    
    hdo = Int(h1 / 3600)
    hfun = Int((h1 - hdo * 3600#) / 60#)
    hbyo = Int(h1# - hdo * 3600# - hfun * 60#)
    vdo = Int(v1 / 3600)
    vfun = Int((v1 - vdo * 3600#) / 60#)
    vbyo = Int(v1 - vdo * 3600# - vfun * 60#)
    Print #3, Format(Time$, "hh:mm:ss"); ","; Format(Pnum, ""); ",";
    Print #3, Chr$(34); rl; Chr$(34); ",";
    Print #3, Format(XF, "0000.000"); ","; Format(YF, "0000.000"); ","; Format(ZF, "0000.000"); ",";
    Print #3, Format(s1, "0000.000"); ",";
    Print #3, Format(hdo, "000000"); ","; Format(hfun, "000"); ","; Format(hbyo, "000"); ",",
    Print #3, Format(vdo, "000000"); ","; Format(vfun, "000"); ","; Format(vbyo, "000")
End Sub

Function senkaiWAIT() As Integer
'�����̈����҂�
' 0:NG
'-1:OK
    Dim h1 As Double, v1 As Double, s1 As Double, c As Integer, rl As String
    Dim ccn As Integer
    senkaiWAIT = 0
    ccn = 0
    c = 4
    While c <> 0 And c <> 1
        ccn = ccn + 1: If ccn > 20 Then Exit Function
        Call DataIn(h1, v1, s1, c, rl)
    Wend
    senkaiWAIT = -1
End Function

Sub BYOtoDMS(byo As Double, dms As String)
'�b����x���b�ɕϊ�
    Dim sg1 As String, d01 As String, fu1 As String, by1 As String
    Dim sg As Integer, d0 As Integer, fu As Integer, bb As Integer
    
    sg = Sgn(byo)
    byo = Abs(byo)
    d0 = Int(byo / 3600#)
    fu = Int((byo - d0 * 3600#) / 60#)
    bb = Int((byo - d0 * 3600# - fu * 60#) * 10#)
    If sg < 0 Then sg1 = "-" Else sg1 = "+"
    d01 = Right$("000" & Right$(str$(d0), Len(str$(d0)) - 1), 3)
    fu1 = Right$("00" & Right$(str$(fu), Len(str$(fu)) - 1), 2)
    by1 = Right$("000" & Right$(str$(bb), Len(str$(bb)) - 1), 3)
    dms = sg1 & d01 & fu1 & by1
End Sub

Sub XyzCal(i As Integer, h1 As Double, v1 As Double, s1 As Double, rl As String)
'�p�x�A�΋���������W�ɕϊ�
    Dim Vdeg As Double, Vrad As Double, Hdeg As Double, Hrad As Double, jdst As Double, xydst As Double
    
    If (rl = "r") Then
        Vdeg = v1 / 3600#
        Vrad = Vdeg * RAD#
        Hdeg = h1 / 3600#
        Hrad = (Hdeg + InitDT.AZIMUTH) * RAD
        jdst = s1
        ZD(i) = jdst * Cos(Vrad) + InitDT.z0 + InitDT.MH
        xydst = jdst * Sin(Vrad)
        XD(i) = xydst * Cos(Hrad) + InitDT.x0
        YD(i) = xydst * Sin(Hrad) + InitDT.y0
    Else
        Vdeg = 360# - (v1 / 3600#)
        Vrad = Vdeg * RAD
        Hdeg = h1 / 3600# + 180#
        Hrad = (Hdeg + InitDT.AZIMUTH) * RAD
        jdst = s1
        ZD(i + InitDT.HeikinKaisuu) = jdst * Cos(Vrad) + InitDT.z0 + InitDT.MH
        xydst = jdst * Sin(Vrad)
        XD(i + InitDT.HeikinKaisuu) = xydst * Cos(Hrad) + InitDT.x0
        YD(i + InitDT.HeikinKaisuu) = xydst * Sin(Hrad) + InitDT.y0
    End If
End Sub

Sub DataCal(Pnum As Integer)
'���E���̃f�[�^�𕽋ς����W���v�Z����
    Dim i As Integer, Hei As Double
    Hei = 2# * InitDT.HeikinKaisuu
    XN = 0: YN = 0: ZN = 0
    For i% = 1 To Hei
        XN = XN + XD(i%)
        YN = YN + YD(i%)
        ZN = ZN + ZD(i%)
    Next
    XN = XN / Hei: YN = YN / Hei: ZN = ZN / Hei
    dx = XN - xo(Pnum): dy = YN - yo(Pnum): dz = ZN - zo(Pnum)
    '''If (iCount = 1) Then
        dx = 0: dy = 0: dz = 0
        xo(Pnum) = XN: yo(Pnum) = YN: zo(Pnum) = ZN
    '''End If
End Sub

Sub TiltIn(h1 As Double, v1 As Double)
    Dim q As Integer, sData As String, rc As Integer
    Dim h2 As String, v2 As String
    Dim dms As String, byo As Double
    
    Call SendCmd("C31")       '�p�x�f�[�^�v��
    rc = ComLinput(sData)
    q% = InStr(sData, "Q")

    h2 = Mid$(sData, q + 9, 9)
    v2 = Mid$(sData, q + 1, 8)
    dms = h2
    Call DMStoBYO(byo, dms)
    h1 = byo
    dms = " " + v2
    Call DMStoBYO(byo, dms)
    v1 = byo
End Sub

Sub ZahyoWrite(FileName As String)
'���W�f�[�^���L�^
    Dim f As Integer, i As Integer
    f = FreeFile
    Open FileName For Append Lock Write As #f
    Print #f, Format(MDY, "YYYY/MM/DD hh:mm:ss");
    For i = 1 To (MAX_CH / 3) 'Tensu
        Print #f, Right$("    " + Format(xo(i), "+####0.000;-####0.000"), 10);
        Print #f, Right$("    " + Format(yo(i), "+####0.000;-####0.000"), 10);
        Print #f, Right$("    " + Format(zo(i), "+####0.000;-####0.000"), 10);
    Next i
    Print #f, ""
    Close (f)
    
    f = FreeFile
    Open CurrentDIR & "final.dat" For Output As #f
    Print #f, Format(MDY, "YYYY/MM/DD hh:mm:ss");
    For i = 1 To (MAX_CH / 3) 'Tensu
        Print #f, Right$("    " + Format(xo(i), "+####0.000;-####0.000"), 10);
        Print #f, Right$("    " + Format(yo(i), "+####0.000;-####0.000"), 10);
        Print #f, Right$("    " + Format(zo(i), "+####0.000;-####0.000"), 10);
    Next i
    Print #f, ""
    Close (f)
    
'    Form2.Text1.SelText = Format(MDY, "YYYY/MM/DD hh:mm:ss")
'    For i = 1 To InitDT.Tensu
'        Form2.Text1.SelText = Right$("    " + Format(xo(i), "+####0.000;-####0.000"), 10)
'        Form2.Text1.SelText = Right$("    " + Format(yo(i), "+####0.000;-####0.000"), 10)
'        Form2.Text1.SelText = Right$("    " + Format(zo(i), "+####0.000;-####0.000"), 10)
'    Next i
'    Form2.Text1.SelText = vbCrLf
End Sub

Sub HouiWrite()
'�����f�[�^���L�^
    Dim f As Integer, i As Integer
    Dim ss As String
    
'    f = FreeFile
'    Open "Houkou.dat" For Output As #f
'    For i = 1 To Tensu
'        Print #f, Right$("    " + Format(H(2, i), "#######0"), 8); ",";
'        Print #f, Right$("    " + Format(V(2, i), "#######0"), 8); ",";
'        Print #f, Right$("    " + Format(S(2, i), "#######0"), 8)
'    Next i
'    Close (f)
    f = FreeFile
    Open InitDT.PoFILE1 For Output Lock Write As #f
    For i = 1 To InitDT.Tensu
        ss = Format(i, "@@@@")
        ss = ss & Right$("            " + Format(PoDT.H(2, i), "#######0"), 12)
        ss = ss & Right$("            " + Format(PoDT.V(2, i), "#######0"), 12)
        ss = ss & Right$("            " + Format(PoDT.s(2, i), "###0.000"), 12)
        Print #f, ss
    Next i
    Close #f

    
'    Form2.Text2.SelText = Format(MDY, "YYYY/MM/DD hh:mm:ss")
'    For i = 1 To InitDT.Tensu
'        Form2.Text2.SelText = Right$("    " + Format(PoDT.H(2, i), "#######0"), 8)
'        Form2.Text2.SelText = Right$("    " + Format(PoDT.V(2, i), "#######0"), 8)
'        Form2.Text2.SelText = Right$("    " + Format(PoDT.s(2, i), "###0.000"), 8)
'    Next i
'    Form2.Text2.SelText = vbCrLf

    For i = 1 To InitDT.Tensu
        PoDT.H(1, i) = PoDT.H(2, i)
        PoDT.V(1, i) = PoDT.V(2, i)
        PoDT.s(1, i) = PoDT.s(2, i)
    Next i
End Sub

Sub SecWait(setS%)
'�w��b���ҋ@
    Dim SecS As Date
    
    SecS = Now
    Do
        DoEvents
        If DateDiff("s", SecS, Now) > setS% Then
            Exit Do
        End If
    Loop
End Sub

Function GTS8on() As Integer
'GTS-8 SW ON
    Dim cmd As String, BC$
    Dim ic As Date
    Dim Srbuf As String
    Dim rc As Integer, RT As String
    
    GTS8on = 0  'NG
    cmd = "G8SW1"
    'BCC�v�Z
    Call BCCcal(cmd, BC$)
    Srbuf = cmd & BC & Chr(EXT) & vbCrLf
    '���M
    rc = RsctlFrm.VBMCom1.SendString(Srbuf)
    'Wait
    Call SecWait(10)
    ic = Now
    Do
        DoEvents
        If DateDiff("s", ic, Now) > 20 Then
            'Stop
'            Call WriteEvents("GTS not WakeUp !!")
            Exit Function
        End If
        Srbuf = cmd & BC & Chr(EXT) & vbCrLf
        '���M
        rc = RsctlFrm.VBMCom1.SendString(Srbuf)
        'ACK��M
        rc = ComLinput(RT$)
    
        If RT$ = Chr(ACK) & "006" & Chr(EXT) Then
            Exit Do
        End If
    Loop
    GTS8on = -1  'OK
End Function

Sub GTS8off()
    Call SendCmd("G8SW0")
End Sub





