Attribute VB_Name = "keijiSet"
Option Explicit

Dim Kkou(Maxkou) As kkou1
Dim dan_ID As Integer, kou_ID As Integer, ten_ID As Integer, type_ID As Integer

'描画変数
Dim PX As Single, PY As Single
Dim PX1 As Single, PY1 As Single, PX2 As Single, PY2 As Single
Dim PANK As String
Dim PANKSIZE As Integer, PANKWIDTH As Integer
Dim ss As String
Dim SIZ As Integer, XOFF As Integer, YOFF As Integer
Dim PENC As Integer
Dim PANKFM As String
Dim SENC As Integer
Dim MKBET As Single
Dim CSA As Integer, CEA As Integer
Dim MD As Double

Public Sub KeijiInit()
    Dim i As Integer, j As Integer, f As Integer, no As Integer
    Dim FLDno As Integer
    Dim L As String
    Dim da As Date
    Dim pdSNG As Single, pdDBL As Double, pdLNG As Long
    Dim po  As Long
    Dim Thistime As Date
    Dim Xbtn As Single
    Dim FileName As String
    Dim dt2(50) As Double         '計測値
    Dim Kpd As Double
    
    'Ｘ軸設定
    Xbtn = (Kkeiji.Xmax * 24) / Kkeiji.xBUN
    Thistime = Now
    Thistime = DateSerial(Year(Thistime), Month(Thistime), day(Thistime)) + 1 '& TimeSerial(Hour(Thistime), 0, 0)
    Kkeiji.ED = Thistime
    Do
        Kkeiji.ED = DateAdd("n", Xbtn * 60, Kkeiji.ED)
        If DateDiff("s", Now, Kkeiji.ED) > 0 Then Exit Do
    Loop
    Kkeiji.SD = DateAdd("h", -(Kkeiji.Xmax * 24), Kkeiji.ED)

Erase Mdt
    
    j = 1
    dan_ID = 1
    kou_ID = 1
    type_ID = 1
    
    KeijiCo(j) = 0
    
    FileName = KEISOKU.Data_path & DATA_DAT
    i = FileCheck(FileName, "計測データ")
    If i = 0 Then Unload frmKeiji1: End

    po = STARTpoint(Kkeiji.SD)
    f = FreeFile
    Open FileName For Input Access Read Shared As #f
        Seek #f, po
        Do While Not (EOF(f))

            Line Input #f, L
            da = CDate(Mid$(L, 1, 19))
            If DateDiff("s", da, Kkeiji.SD) > 0 Then GoTo Kskip1
            If DateDiff("s", Kkeiji.ED, da) > 0 Then Exit Do
            If DateDiff("s", Now, da) > 0 Then Exit Do
            
            KeijiCo(j) = KeijiCo(j) + 1
            Mdt(j, KeijiCo(j)).day = da
            
            For i = 1 To 2
                If i = 1 Then ten_ID = 3
                If i = 2 Then ten_ID = 6
                FLDno = Tbl(kou_ID, dan_ID, ten_ID).FLD
                    
                If IsNumeric(Mid$(L, 20 + 10 * (FLDno - 1), 10)) = True Then
                    pdDBL = CDbl(Mid$(L, 20 + 10 * (FLDno - 1), 10))
                Else
                    pdDBL = 999999
                End If

                If Abs(pdDBL) >= 999999 Or Tbl(kou_ID, dan_ID, ten_ID).Syo = 999999 Then
                    dt2(ten_ID) = 999999
                Else
                    '2001/12/11
                    dt2(ten_ID) = (pdDBL - Tbl(kou_ID, dan_ID, ten_ID).Syo) * Tbl(kou_ID, dan_ID, ten_ID).Kei
                    ''dt2(ten_ID) = (-1) * (pdDBL - Tbl(kou_ID, dan_ID, ten_ID).Syo) * Tbl(kou_ID, dan_ID, ten_ID).Kei
                End If
            Next i
            Call KEISAN(dt2(), Kpd)
            Mdt(j, KeijiCo(j)).data(1) = Kpd
Kskip1:
        Loop
    Close
    
    Call KeijiPlot1
End Sub

Public Sub KeijiPlot1()
    Dim i As Integer, j As Integer, co As Integer
    Dim Thistime As Date
    Dim Xbtn As Single, Ybtn As Single
    Dim YScl As Single
    Dim FLDno As Integer
    Dim da As Date
    Dim pdSNG As Single, pdDBL As Double
    Dim HI As Date
    Dim SW As Integer ', Ksw As Boolean
    Dim YMIN As Single, YMAX As Single, yBUN As Integer
    Dim Kpd As Double
    Dim Kmin As Single, Kmax As Single
    Dim SX As Single, SY As Single, Ex As Single, Ey As Single
    
    frmKeiji1.VSDraw1.Visible = False
    frmKeiji1.VSDraw1.Clear
    PENC = 1: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
    SENC = 0: Call VD_LTCD(frmKeiji1.VSDraw1, SENC)
    
    Kkeiji.XP = 8500
    Kkeiji.YP = 8500
    Kkeiji.XS = 900
    Kkeiji.YS = 9500
    
    Kkeiji.XP = 8000
    Kkeiji.YP = 8000
    Kkeiji.XS = 900
    Kkeiji.YS = 8800
    
    Kkeiji.XP = 17000
    Kkeiji.YP = 17000 \ 2
    Kkeiji.XS = 2000
    Kkeiji.YS = 19000 \ 2
    
    '軸設定
    Xbtn = (Kkeiji.Xmax * 24) / Kkeiji.xBUN
    Kkeiji.xb = Kkeiji.XP / (Kkeiji.Xmax * 24)
    dan_ID = 1
    kou_ID = 1
    type_ID = 1
    
    YMIN = Kkeiji.YMIN
    YMAX = Kkeiji.YMAX
    yBUN = Kkeiji.yBUN
    
    Kkou(1).yb = Kkeiji.YP / (YMAX - YMIN)
    Ybtn = (YMAX - YMIN) / yBUN
    
    PX1 = Kkeiji.XS: PX2 = Kkeiji.XS + Kkeiji.XP
    PY2 = Kkeiji.YS '- (i - 1) * Kkeiji.YP
    PY1 = PY2 - Kkeiji.YP
    
    '管理レベル色塗りつぶし
    frmKeiji1.VSDraw1.PenWidth = 0
    frmKeiji1.VSDraw1.PenColor = QBColor(7)
    For i = 1 To 4
        If Kanri(kou_ID, dan_ID).Lebel2(i) > YMIN And Kanri(kou_ID, dan_ID).Lebel1(i) < YMAX Then
            If Kanri(kou_ID, dan_ID).Lebel1(i) < YMIN Then Kmin = YMIN Else Kmin = Kanri(kou_ID, dan_ID).Lebel1(i)
            If Kanri(kou_ID, dan_ID).Lebel2(i) > YMAX Then Kmax = YMAX Else Kmax = Kanri(kou_ID, dan_ID).Lebel2(i)
            SX = Kkeiji.XS: Ex = Kkeiji.XS + Kkeiji.XP
            SY = PY2 - Kkou(1).yb * (Kmin - YMIN): Ey = PY2 - Kkou(1).yb * (Kmax - YMIN)
            If i = 1 Then frmKeiji1.VSDraw1.BrushColor = &H80FF80
            If i = 2 Then frmKeiji1.VSDraw1.BrushColor = &H80FFFF
            If i = 3 Then frmKeiji1.VSDraw1.BrushColor = &HFF80FF
            If i = 4 Then frmKeiji1.VSDraw1.BrushColor = RGB(256, 60, 60)
            
            Call VD_BRectangle(frmKeiji1.VSDraw1, SX, SY, Ex, Ey, 0)
        End If
    Next i
    
    '縦軸
    frmKeiji1.VSDraw1.PenWidth = 0
    PENC = 1: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
'    PANKSIZE = 180: Call VD_AnkCsize(frmKeiji1.VSDraw1, PANKSIZE)
    Thistime = Kkeiji.SD
    For i = 0 To Kkeiji.xBUN
        frmKeiji1.VSDraw1.PenColor = QBColor(0)
        PX = Kkeiji.XS + i * (Kkeiji.XP / Kkeiji.xBUN): PY = (Kkeiji.YS - Kkeiji.YP): Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
        PY = PY - 100: Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        PX = Kkeiji.XS + i * (Kkeiji.XP / Kkeiji.xBUN) - 300 - 700
        PY = (Kkeiji.YS - Kkeiji.YP) - 150 - 100
        ss = Format(Format$(Thistime, "yy/m/d"), "@@@@@@@@")
        SIZ = 180 * 2: XOFF = 180 * 2: YOFF = 0: Call VD_KPUT(frmKeiji1.VSDraw1, ss, PX, PY, SIZ, XOFF, YOFF, 0, 1) 'call VD_PPANK(frmKeiji1.VSDraw1, PANK, PX, PY)
'        PY = (Kkeiji.YS - Kkeiji.YP) - 350: PANK = Format$(Thistime, "hh:nn"): Call VD_PPANK(frmKeiji1.VSDraw1, PANK, PX, PY)
        Thistime = DateAdd("n", Xbtn * 60, Thistime)
        If i > 0 And i < Kkeiji.xBUN Then
            frmKeiji1.VSDraw1.PenColor = QBColor(7)
            PX = Kkeiji.XS + i * (Kkeiji.XP / Kkeiji.xBUN)
            PY = (Kkeiji.YS - Kkeiji.YP): Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
            PY = PY + Kkeiji.YP: Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        End If
    Next i
    
    '横軸
    PANKSIZE = 160 + 200: Call VD_AnkCsize(frmKeiji1.VSDraw1, PANKSIZE)
    YScl = YMIN
    For j = 0 To yBUN
        
        '軸
        frmKeiji1.VSDraw1.PenWidth = 0
        PY = PY2 - j * (Kkeiji.YP / yBUN)
        PENC = 1: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
        PX = Kkeiji.XS: Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
        PX = Kkeiji.XS - 50: Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        
        If j = 0 Then
            frmKeiji1.VSDraw1.PenColor = QBColor(0)
            frmKeiji1.VSDraw1.PenWidth = 30
        Else
            frmKeiji1.VSDraw1.PenColor = QBColor(7)
            frmKeiji1.VSDraw1.PenWidth = 0
        End If
        
        PX = Kkeiji.XS: Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
        PX = Kkeiji.XS + Kkeiji.XP: Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        
        'スケール
        PY = PY + 100 - 40 + 100
        PX = Kkeiji.XS - 600 - 700
        PANK = Format$(Format(YScl, "0.0"), "@@@@@@@")
        Call VD_PPANK(frmKeiji1.VSDraw1, PANK, PX, PY)
        
        YScl = YScl + Ybtn
    Next j
    
    
    '枠
    PENC = 1: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
    frmKeiji1.VSDraw1.PenWidth = 30
    Call VD_MMM(frmKeiji1.VSDraw1, PX1, PY1)
    Call VD_DDD(frmKeiji1.VSDraw1, PX1, PY2)
    Call VD_DDD(frmKeiji1.VSDraw1, PX2, PY2)
    Call VD_DDD(frmKeiji1.VSDraw1, PX2, PY1)
    Call VD_DDD(frmKeiji1.VSDraw1, PX1, PY1)
    frmKeiji1.VSDraw1.PenWidth = 0
    
    
    '管理レベルテキスト
    For i = 1 To 4
        If Kanri(kou_ID, dan_ID).Lebel2(i) > YMIN And Kanri(kou_ID, dan_ID).Lebel1(i) < YMAX Then
            If Kanri(kou_ID, dan_ID).Lebel1(i) < YMIN Then Kmin = YMIN Else Kmin = Kanri(kou_ID, dan_ID).Lebel1(i)
            If Kanri(kou_ID, dan_ID).Lebel2(i) > YMAX Then Kmax = YMAX Else Kmax = Kanri(kou_ID, dan_ID).Lebel2(i)
            ss = Trim(Kanri(kou_ID, dan_ID).TI1(i)) & "（" & Trim(Kanri(kou_ID, dan_ID).TI2(i)) & "）"
            SIZ = 152 * 2: XOFF = 180 * 2: YOFF = 0
            PX = Kkeiji.XS + 50
            PY = PY2 - Kkou(1).yb * ((Kmin + (Kmax - Kmin) / 2) - YMIN) + 100
            Call VD_KPUT(frmKeiji1.VSDraw1, ss, PX, PY, SIZ, XOFF, YOFF, 0, 0)
        End If
    Next i
    
    '物理量
    ss = Trim(Kkeiji.Yti): j = LenB(StrConv(ss, vbFromUnicode))
    PX = PX1 - 800 + 100 - 800: PY = PY2 - Kkeiji.YP / 2 + j * 100: SIZ = 152 * 2: XOFF = 0: YOFF = -200 * 2
    Call VD_KPUT(frmKeiji1.VSDraw1, ss, PX, PY, SIZ, XOFF, YOFF, 0, 1)
    
    ss = Trim(Kkeiji.Yu)
    PX = PX1 - 800 - 400 + 100 - 800: PY = PY2 - Kkeiji.YP / 2 - 80 - j * 100: SIZ = 152 * 2: XOFF = 200 * 2: YOFF = 0
    Call VD_KPUT(frmKeiji1.VSDraw1, ss, PX, PY, SIZ, XOFF, YOFF, 0, 1)
    
    
    i = 1
    dan_ID = 1
    kou_ID = 1
    type_ID = 1

    YMIN = Kkeiji.YMIN
    YMAX = Kkeiji.YMAX
    
    Kkou(1).yb = Kkeiji.YP / (YMAX - YMIN)

    PX1 = Kkeiji.XS: PX2 = Kkeiji.XS + Kkeiji.XP
    PY2 = Kkeiji.YS - (i - 1) * Kkeiji.YP
    PY1 = PY2 - Kkeiji.YP

    
    j = 1
    FLDno = Tbl(kou_ID, dan_ID, Kkou(1).ten(j)).FLD

    PENC = j: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
    frmKeiji1.VSDraw1.BrushStyle = bsTransparent

    SW = 0
    For co = 1 To KeijiCo(i)
        da = Mdt(i, co).day
        If DateDiff("s", da, Kkeiji.SD) > 0 Then GoTo Kskip2
        If DateDiff("s", Kkeiji.ED, da) > 0 Then Exit For

        HI = da
        MD = DateDiff("s", Kkeiji.SD, HI) / 3600 '86400#

        pdDBL = Mdt(i, co).data(j)
        
        If pdDBL >= 999999 Then SW = 0: GoTo Kskip2
        
        If (PY2 - Kkou(1).yb * (pdDBL - YMIN)) < PY1 Then SW = 0: GoTo Kskip2
        If (PY2 - Kkou(1).yb * (pdDBL - YMIN)) > PY1 + Kkeiji.YP Then SW = 0: GoTo Kskip2
        
        Select Case SW
            Case 0
                PY = PY2 - Kkou(1).yb * (pdDBL - YMIN)
                PX = Kkeiji.XS + MD * Kkeiji.xb
                Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
                SW = 1
            Case 1
                PY = PY2 - Kkou(1).yb * (pdDBL - YMIN)
                PX = Kkeiji.XS + MD * Kkeiji.xb
                Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        End Select
        Call VD_MK(frmKeiji1.VSDraw1, PX, PY, 1)
        
Kskip2:

    Next co

    
    frmKeiji1.VSDraw1.Show
    frmKeiji1.VSDraw1.Visible = True
End Sub

Public Sub KeijiPlot2()
    Dim i As Integer, j As Integer
    Dim pd1 As Double, pd2 As Single, pdLNG As Long, pdSNG As Single, pdDBL As Double
    Dim HI As Date
    Dim no As Integer
    Dim Z_MD As Double
    Dim YMIN As Single
    Dim dt2(50) As Double         '計測値
    Dim Kpd As Double
    
    frmKeiji1.VSDraw1.BrushStyle = bsTransparent
    no = 1
    dan_ID = 1
    kou_ID = 1
    type_ID = 1
    
    KeijiCo(no) = KeijiCo(no) + 1

    HI = Z_Keisoku_Time
    MD = DateDiff("s", Kkeiji.SD, HI) / 3600 '86400#
    Mdt(no, KeijiCo(no)).day = HI
    
    If KeijiCo(no) = 1 Then
        Z_MD = MD
    Else
        Z_MD = DateDiff("s", Kkeiji.SD, Mdt(no, KeijiCo(no) - 1).day) / 3600
    End If

    YMIN = kou(kou_ID, type_ID).Kmin
    
    PX1 = Kkeiji.XS
    PX2 = Kkeiji.XS + Kkeiji.XP
    PY2 = Kkeiji.YS - (no - 1) * Kkeiji.YP
    PY1 = PY2 - Kkeiji.YP
    
    
    For i = 1 To 2
        If i = 1 Then ten_ID = 3
        If i = 2 Then ten_ID = 6
        pdDBL = dt1(Tbl(kou_ID, dan_ID, ten_ID).ch)
        If Abs(pdDBL) >= 999999 Or Tbl(kou_ID, dan_ID, ten_ID).Syo = 999999 Then
            dt2(ten_ID) = 999999
        Else
            '2001/12/11
            dt2(ten_ID) = (pdDBL - Tbl(kou_ID, dan_ID, ten_ID).Syo) * Tbl(kou_ID, dan_ID, ten_ID).Kei
            ''dt2(ten_ID) = (-1) * (Tbl(kou_ID, dan_ID, ten_ID).Syo - pdDBL) * Tbl(kou_ID, dan_ID, ten_ID).Kei
        End If
    Next i
    Call KEISAN(dt2(), Kpd)
    
    j = 1
    pd1 = CDbl(Kpd)
    
    Mdt(no, KeijiCo(no)).data(j) = pd1     '今回データ
    pd2 = Mdt(no, KeijiCo(no) - 1).data(j) '前回データ
    
    PENC = j: Call VD_PENJ(frmKeiji1.VSDraw1, PENC)
    
    If (PY2 - Kkou(1).yb * (pd1 - YMIN)) < PY1 Then GoTo Kskip
    If (PY2 - Kkou(1).yb * (pd1 - YMIN)) > PY1 + Kkeiji.YP Then GoTo Kskip
    
    If Mdt(no, KeijiCo(no) - 1).data(j) = 999999 Or KeijiCo(no) = 1 Then
        PY = PY2 - Kkou(1).yb * (pd1 - YMIN)
        PX = Kkeiji.XS + MD * Kkeiji.xb
        Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
        Call VD_MK(frmKeiji1.VSDraw1, PX, PY, 1)
    Else
        PY = PY2 - Kkou(1).yb * (pd2 - YMIN)
        PX = Kkeiji.XS + Z_MD * Kkeiji.xb
        Call VD_MMM(frmKeiji1.VSDraw1, PX, PY)
        
        PY = PY2 - Kkou(1).yb * (pd1 - YMIN)
        PX = Kkeiji.XS + MD * Kkeiji.xb
        Call VD_DDD(frmKeiji1.VSDraw1, PX, PY)
        Call VD_MK(frmKeiji1.VSDraw1, PX, PY, 1)
    End If
Kskip:
    
    frmKeiji1.VSDraw1.Show
End Sub

