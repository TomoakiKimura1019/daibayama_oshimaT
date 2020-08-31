VERSION 5.00
Object = "{7E00A3A2-8F5C-11D2-BAA4-04F205C10000}#1.0#0"; "VSVIEW6.ocx"
Begin VB.Form frmKeiji 
   Caption         =   "経時変化図"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7995
   Icon            =   "frmKeiji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows の既定値
   Begin VSVIEW6Ctl.VSDraw VSDraw1 
      Height          =   8000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8000
      _cx             =   14111
      _cy             =   14111
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
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
   End
   Begin VB.Menu mnuSet 
      Caption         =   "設定"
   End
End
Attribute VB_Name = "frmKeiji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Kkou(Maxkou) As kkou1
Dim dan_ID As Integer, kou_ID As Integer, ten_ID As Integer, type_ID As Integer

'描画変数
Dim PX As Single, PY As Single
Dim PX1 As Single, PY1 As Single, PX2 As Single, PY2 As Single
Dim PANK As String
Dim PANKSIZE As Integer, PANKWIDTH As Integer
Dim SS As String
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
    Dim FILENAME As String
    Dim dt2(50) As Double         '計測値

    i = FileCheck(CurrentDIR & "keiji.dat", "経時変化図設定データ")
    If i = 0 Then Unload 計測Form: End

    Erase Kkou
    
    f = FreeFile: no = 0
    Open CurrentDIR & "keiji.dat" For Input As #f
        For i = 1 To 4
            Input #f, L: Input #f, L
            j = CInt(L)
            If j = 1 Then
                no = no + 1
                Input #f, L: Kkou(no).dan = CInt(L)
                Input #f, L: Kkou(no).kou = CInt(L)
                Input #f, L: Kkou(no).type = CInt(L)
                Input #f, L: Kkou(no).ten(0) = CInt(L)
                For j = 1 To Kkou(no).ten(0)
                    Input #f, L
                    Kkou(no).ten(j) = CInt(L)
                Next j
            End If
        Next i
    Close #f
    Kkou(0).dan = no
    
    'Ｘ軸設定
    Xbtn = keiji.XMAX / keiji.Xbun
    Thistime = Now
    Thistime = DateSerial(Year(Thistime), Month(Thistime), day(Thistime)) & TimeSerial(Hour(Thistime), 0, 0)
    keiji.ed = Thistime
    Do
        keiji.ed = DateAdd("n", Xbtn * 60, keiji.ed)
        If DateDiff("s", Now, keiji.ed) > 0 Then Exit Do
    Loop
    keiji.sd = DateAdd("h", -(keiji.XMAX), keiji.ed)

Erase Mdt
    For j = 1 To Kkou(0).dan
        dan_ID = Kkou(j).dan
        kou_ID = Kkou(j).kou
        type_ID = 1
        
        KeijiCo(j) = 0
        
        FILENAME = KEISOKU.Data_path & DATA_DAT
        i = FileCheck(FILENAME, "計測データ")
        If i = 0 Then Unload 計測Form: End

        po = STARTpoint(FILENAME, kou_ID, keiji.sd)
        f = FreeFile
        Open FILENAME For Input Shared As #f
            Seek #f, po
            Do While Not (EOF(f))
    
                Line Input #f, L
                da = CDate(Mid$(L, 1, 19))
                If DateDiff("s", da, keiji.sd) > 0 Then GoTo Kskip1
                If DateDiff("s", keiji.ed, da) > 0 Then Exit Do
                If DateDiff("s", Now, da) > 0 Then Exit Do
                
                KeijiCo(j) = KeijiCo(j) + 1
                Mdt(j, KeijiCo(j)).day = da
                
                For ten_ID = 1 To Tbl(kou_ID, dan_ID, 0).ten
                    FLDno = Tbl(kou_ID, dan_ID, ten_ID).FLD
                    
                    If IsNumeric(Mid$(L, 20 + 8 * (FLDno - 1), 8)) = True Then
                        pdSNG = CSng(Mid$(L, 20 + 8 * (FLDno - 1), 8))
                    Else
                        pdSNG = 999999
                    End If

                    If Abs(pdSNG) >= 999999 Or Tbl(kou_ID, dan_ID, ten_ID).Syo = 999999 Then
                        dt2(ten_ID) = 999999
                    Else
                        If kou_ID = 1 Then Call SinsyukuCAL(dan_ID, ten_ID, da, pdSNG)
                        dt2(ten_ID) = (pdSNG - Tbl(kou_ID, dan_ID, ten_ID).Syo) * Tbl(kou_ID, dan_ID, ten_ID).Kei
                    End If
'''                    dt2(ten_ID) = pdDBL
                Next ten_ID
'''                Call KEISAN(da, kou_ID, dan_ID, type_ID, dt2())
                
                For ten_ID = 1 To Kkou(j).ten(0)
                    Mdt(j, KeijiCo(j)).data(ten_ID) = CSng(dt2(Kkou(j).ten(ten_ID)))
                Next ten_ID
Kskip1:
            Loop
        Close
    Next j
    
    Call KeijiPlot1
End Sub

Public Sub KeijiPlot1()
    Dim i As Integer, j As Integer, co As Integer
    Dim Thistime As Date
    Dim Xbtn As Single, Ybtn As Single
    Dim yscl As Single
    Dim FLDno As Integer
    Dim da As Date
    Dim pdSNG As Single, pdDBL As Double
    Dim HI As Date
    Dim SW As Integer ', Ksw As Boolean
    Dim YMIN As Single, YMAX As Single, YBUN As Integer

    VSDraw1.Visible = False
    VSDraw1.Clear
    PENC = 1: Call PENJ(VSDraw1, PENC)
    SENC = 0: Call LTCD(VSDraw1, SENC)
    
    keiji.XP = 7000
    keiji.YP = (700 \ Kkou(0).dan) * 10
    keiji.XS = 700
    keiji.YS = 7400
    
    'Ｘ軸設定
    Xbtn = keiji.XMAX / keiji.Xbun
    keiji.Xb = keiji.XP / keiji.XMAX
    
    PANKSIZE = 180: Call AnkCsize(VSDraw1, PANKSIZE)
    
    Thistime = keiji.sd
    For i = 0 To keiji.Xbun
        PX = keiji.XS + i * (keiji.XP / keiji.Xbun): PY = keiji.YS: Call MMM(VSDraw1, PX, PY)
        PY = PY + 50: Call DDD(VSDraw1, PX, PY)
        PX = keiji.XS + i * (keiji.XP / keiji.Xbun) - 300
        PY = keiji.YS + 500: PANK = Format$(Thistime, "mm/dd"): Call PPANK(VSDraw1, PANK, PX, PY)
        PY = keiji.YS + 250: PANK = Format$(Thistime, "hh:nn"): Call PPANK(VSDraw1, PANK, PX, PY)
        Thistime = DateAdd("n", Xbtn * 60, Thistime)
        If i > 0 And i < keiji.Xbun Then
            VSDraw1.PenColor = QBColor(7)
            PX = keiji.XS + i * (keiji.XP / keiji.Xbun)
            PY = keiji.YS: Call MMM(VSDraw1, PX, PY)
            PY = PY - keiji.YP * Kkou(0).dan: Call DDD(VSDraw1, PX, PY)
        End If
    Next i
    
    For i = 1 To Kkou(0).dan
    
        dan_ID = Kkou(i).dan
        kou_ID = Kkou(i).kou
        type_ID = 1
        
        YMIN = kou(kou_ID, type_ID).KMIN
        YMAX = kou(kou_ID, type_ID).KMAX
        YBUN = kou(kou_ID, type_ID).KBUN
        
        Kkou(i).YB = keiji.YP / (YMAX - YMIN)
        Ybtn = (YMAX - YMIN) / YBUN
        
        PENC = 1: Call PENJ(VSDraw1, PENC)
        PX1 = keiji.XS: PX2 = keiji.XS + keiji.XP
        PY2 = keiji.YS - (i - 1) * keiji.YP
        PY1 = PY2 - keiji.YP
        
        VSDraw1.PenWidth = 30
        Call MMM(VSDraw1, PX1, PY1)
        Call DDD(VSDraw1, PX1, PY2)
        Call DDD(VSDraw1, PX2, PY2)
        Call DDD(VSDraw1, PX2, PY1)
        Call DDD(VSDraw1, PX1, PY1)
        
        
        PANKSIZE = 160: Call AnkCsize(VSDraw1, PANKSIZE)
        yscl = YMIN
        For j = 0 To YBUN - 1
            
            '軸
            VSDraw1.PenWidth = 0
            PY = PY1 + j * (keiji.YP / YBUN)
            PENC = 1: Call PENJ(VSDraw1, PENC)
            PX = keiji.XS: Call MMM(VSDraw1, PX, PY)
            PX = keiji.XS - 50: Call DDD(VSDraw1, PX, PY)
            
            If j = 0 Then
                VSDraw1.PenColor = QBColor(0)
                VSDraw1.PenWidth = 30
            Else
                VSDraw1.PenColor = QBColor(7)
                VSDraw1.PenWidth = 0
            End If
            
            PX = keiji.XS: Call MMM(VSDraw1, PX, PY)
            PX = keiji.XS + keiji.XP: Call DDD(VSDraw1, PX, PY)
            
            'スケール
            PY = PY + 100 - 40: PX = keiji.XS - 600
            PANK = Format$(Format(yscl, "0.0"), "@@@@@@@")
            Call PPANK(VSDraw1, PANK, PX, PY)
            
            yscl = yscl + Ybtn
        Next j
        
        VSDraw1.PenWidth = 0
        PENC = 1: Call PENJ(VSDraw1, PENC)
        
        PY = PY1 + Kkou(i).YB * (0 - YMIN)
        PX = keiji.XS: Call MMM(VSDraw1, PX, PY)
        PX = keiji.XS + keiji.XP: Call DDD(VSDraw1, PX, PY)
        
'''        If KEISOKU.Datatype = 1 Then
'''            Ksw = False
'''            SENC = 2: Call LTCD(VSDraw1, SENC)
'''
'''            If kou_ID = 1 And type_ID = 1 Then Ksw = True '層別沈下累計変位
'''            If kou_ID = 2 And type_ID = 2 Then Ksw = True '歪み計応力度
'''            If kou_ID = 4 And type_ID = 2 Then Ksw = True '無応力計応力度
'''            If kou_ID = 6 And type_ID = 1 Then Ksw = True '土圧
'''            If kou_ID = 7 And type_ID = 1 Then Ksw = True '塩ビ
'''            If Ksw = True Then
'''
'''                For j = 1 To 4
'''                    If j = 1 Then VSDraw1.PenColor = QBColor(14)
'''                    If j = 2 Then VSDraw1.PenColor = &H80FF&
'''                    If j = 3 Then VSDraw1.PenColor = QBColor(13)
'''                    If j = 4 Then VSDraw1.PenColor = QBColor(12)
'''
'''                    If Kanri(kou_ID,dan_ID, type_ID, 1, 1).Lebel(j) <> 0 And Kanri(kou_ID,dan_ID, type_ID, 1, 1).keihouSW = 0 Then GoTo Kskip3
'''                    If Kanri(kou_ID,dan_ID, type_ID, 1, 1).Lebel(j) = 0 Then GoTo Kskip3
'''                    If Abs(Kanri(kou_ID,dan_ID, type_ID, 1, 1).Lebel(j)) = 999999 Then GoTo Kskip3
'''
'''                    If YMAX > Kanri(kou_ID,dan_ID, type_ID, 1, 1).Lebel(j) Then
'''                        PY = PY1 + Kkou(i).YB * (Kanri(kou_ID,dan_ID, type_ID, 1, 1).Lebel(j) - YMIN)
'''                        PX = keiji.XS: Call MMM(VSDraw1, PX, PY)
'''                        PX = keiji.XS + keiji.XP: Call DDD(VSDraw1, PX, PY)
'''
'''                    End If
'''Kskip3:
'''                    If Kanri(kou_ID,dan_ID, type_ID, 2, 1).Lebel(j) <> 0 And Kanri(kou_ID,dan_ID, type_ID, 2, 1).keihouSW = 0 Then GoTo Kskip4
'''                    If Kanri(kou_ID,dan_ID, type_ID, 2, 1).Lebel(j) = 0 Then GoTo Kskip4
'''                    If Abs(Kanri(kou_ID,dan_ID, type_ID, 2, 1).Lebel(j)) = 999999 Then GoTo Kskip4
'''
'''                    If YMIN < Kanri(kou_ID,dan_ID, type_ID, 2, 1).Lebel(j) Then
'''                        PY = PY1 + Kkou(i).YB * (Kanri(kou_ID,dan_ID, type_ID, 2, 1).Lebel(j) - YMIN)
'''                        PX = keiji.XS: Call MMM(VSDraw1, PX, PY)
'''                        PX = keiji.XS + keiji.XP: Call DDD(VSDraw1, PX, PY)
'''                    End If
'''Kskip4:
'''                Next j
'''            End If
'''            SENC = 0: Call LTCD(VSDraw1, SENC)
'''        End If
'''
        PENC = 1: Call PENJ(VSDraw1, PENC)
        
        '物理量
        PX = PX1 + 100: PY = PY2 - 50: SIZ = 152: XOFF = 200: YOFF = 0
        SS = Trim$(DanSet(kou_ID, dan_ID).ti) & " " & Trim$(kou(kou_ID, type_ID).TI1)
        If kou(kou_ID, 0).no > 1 Then SS = SS & " " & Trim$(kou(kou_ID, type_ID).TI2)
        SS = SS & " (" & Trim$(kou(kou_ID, type_ID).Yu) & ")"
        Call KPUT(VSDraw1, SS, PX, PY, SIZ, XOFF, YOFF, 0, 1)
        
        If Tbl(kou_ID, dan_ID, 0).ten > 1 Then
            For j = 1 To Kkou(i).ten(0)
                PENC = j + 1: Call PENJ(VSDraw1, PENC)
                If j > 8 Then
                    PX = PX1 + 2500 + (j - 8) * 1200 - 2000: PY = PY2 - 130 - 360
                ElseIf j > 4 Then
                    PX = PX1 + 2500 + (j - 4) * 1200 - 2000: PY = PY2 - 130 - 180
                Else
                    PX = PX1 + 2500 + j * 1200 - 2000: PY = PY2 - 130
                End If
                Call MMM(VSDraw1, PX, PY)
                PX = PX + 250: Call DDD(VSDraw1, PX, PY)
                
                PENC = 1: Call PENJ(VSDraw1, PENC)
                If j > 8 Then
                    PX = PX1 + 3000 + (j - 8) * 1200 - 2000 - 200: PY = PY2 - 50 - 360
                ElseIf j > 4 Then
                    PX = PX1 + 3000 + (j - 4) * 1200 - 2000 - 200: PY = PY2 - 50 - 180
                Else
                    PX = PX1 + 3000 + j * 1200 - 2000 - 200: PY = PY2 - 50
                End If
                
                SIZ = 152: XOFF = 200: YOFF = 0
                SS = Trim$(Tbl(kou_ID, dan_ID, Kkou(i).ten(j)).HAN)
                Call KPUT(VSDraw1, SS, PX, PY, SIZ, XOFF, YOFF, 0, 1)
            Next j
        End If
    Next i

    For i = 1 To Kkou(0).dan
    
        dan_ID = Kkou(i).dan
        kou_ID = Kkou(i).kou
        type_ID = 1

        YMIN = kou(kou_ID, type_ID).KMIN
        YMAX = kou(kou_ID, type_ID).KMAX
        
        Kkou(i).YB = keiji.YP / (YMAX - YMIN)

        PX1 = keiji.XS: PX2 = keiji.XS + keiji.XP
        PY2 = keiji.YS - (i - 1) * keiji.YP
        PY1 = PY2 - keiji.YP


        For j = 1 To Kkou(i).ten(0)

            FLDno = Tbl(kou_ID, dan_ID, Kkou(i).ten(j)).FLD

            PENC = j + 1: Call PENJ(VSDraw1, PENC)

            SW = 0
            For co = 1 To KeijiCo(i)
                da = Mdt(i, co).day
                If DateDiff("s", da, keiji.sd) > 0 Then GoTo Kskip2
                If DateDiff("s", keiji.ed, da) > 0 Then Exit For
        
                HI = da
                MD = DateDiff("s", keiji.sd, HI) / 3600 '86400#

                pdSNG = Mdt(i, co).data(j)
                
                If pdSNG >= 999999 Then SW = 0: GoTo Kskip2
                
                If (PY1 + Kkou(i).YB * (pdSNG - YMIN)) < PY1 Then SW = 0: GoTo Kskip2
                If (PY1 + Kkou(i).YB * (pdSNG - YMIN)) > PY1 + keiji.YP Then SW = 0: GoTo Kskip2
                
                Select Case SW
                    Case 0
                        PY = PY1 + Kkou(i).YB * (pdSNG - YMIN)
                        PX = keiji.XS + MD * keiji.Xb
                        Call MMM(VSDraw1, PX, PY)
                        SW = 1
                    Case 1
                        PY = PY1 + Kkou(i).YB * (pdSNG - YMIN)
                        PX = keiji.XS + MD * keiji.Xb
                        Call DDD(VSDraw1, PX, PY)
                End Select
    
Kskip2:

            Next co
        Next j
    Next i

    
    VSDraw1.Show
    VSDraw1.Visible = True
End Sub

Public Sub KeijiPlot2()
    Dim i As Integer, j As Integer
    Dim pd1 As Single, pd2 As Single, pdLNG As Long, pdSNG As Single
    Dim HI As Date
    Dim no As Integer
    Dim Z_MD As Double
    Dim YMIN As Single
    Dim dt2(50) As Double         '計測値

    For no = 1 To Kkou(0).dan
    
        dan_ID = Kkou(no).dan
        kou_ID = Kkou(no).kou
        type_ID = 1
        
        KeijiCo(no) = KeijiCo(no) + 1
    
        HI = Z_Keisoku_Time
        MD = DateDiff("s", keiji.sd, HI) / 3600 '86400#
        Mdt(no, KeijiCo(no)).day = HI
        
        If KeijiCo(no) = 1 Then
            Z_MD = MD
        Else
            Z_MD = DateDiff("s", keiji.sd, Mdt(no, KeijiCo(no) - 1).day) / 3600
        End If
    
        YMIN = kou(kou_ID, type_ID).KMIN
        
        PX1 = keiji.XS
        PX2 = keiji.XS + keiji.XP
        PY2 = keiji.YS - (no - 1) * keiji.YP
        PY1 = PY2 - keiji.YP
        
        For ten_ID = 1 To Tbl(kou_ID, dan_ID, 0).ten
            pdSNG = dt1(Tbl(kou_ID, dan_ID, ten_ID).ch)
            If Abs(pdSNG) >= 999999 Or Tbl(kou_ID, dan_ID, ten_ID).Syo = 999999 Then
                dt2(ten_ID) = 999999
            Else
                If kou_ID = 1 Then Call SinsyukuCAL(dan_ID, ten_ID, HI, pdSNG)
                dt2(ten_ID) = (pdSNG - Tbl(kou_ID, dan_ID, ten_ID).Syo) * Tbl(kou_ID, dan_ID, ten_ID).Kei
            End If
        Next ten_ID
'''        Call KEISAN(HI, kou_ID, dan_ID, type_ID, dt2())
        
        For j = 1 To Kkou(no).ten(0)
            
            pd1 = CSng(dt2(Kkou(no).ten(j)))
            
            Mdt(no, KeijiCo(no)).data(j) = pd1     '今回データ
            pd2 = Mdt(no, KeijiCo(no) - 1).data(j) '前回データ
            
            PENC = j + 1: Call PENJ(VSDraw1, PENC)
            
            If (PY1 + Kkou(no).YB * (pd1 - YMIN)) < PY1 Then GoTo Kskip
            If (PY1 + Kkou(no).YB * (pd1 - YMIN)) > PY1 + keiji.YP Then GoTo Kskip
            
            If Mdt(no, KeijiCo(no) - 1).data(j) = 999999 Or KeijiCo(no) = 1 Then
                PY = PY1 + Kkou(no).YB * (pd1 - YMIN)
                PX = keiji.XS + MD * keiji.Xb
                Call MMM(VSDraw1, PX, PY)
            Else
                PY = PY1 + Kkou(no).YB * (pd2 - YMIN)
                PX = keiji.XS + Z_MD * keiji.Xb
                Call MMM(VSDraw1, PX, PY)
                
                PY = PY1 + Kkou(no).YB * (pd1 - YMIN)
                PX = keiji.XS + MD * keiji.Xb
                Call DDD(VSDraw1, PX, PY)
            End If
Kskip:
        Next j
    
    Next no
    
    VSDraw1.Show
End Sub

Private Sub Form_Load()
    Left = Screen.Width - Me.Width
    top = 0
    
    'グラフ
    VSDraw1.PenWidth = 1           ' 線幅
    VSDraw1.FontName = "ＭＳ ゴシック"
    VSDraw1.BrushStyle = bsTransparent
    VSDraw1.ScaleLeft = 0
    VSDraw1.ScaleWidth = 8000 '11500
    VSDraw1.ScaleTop = 8000 '19530
    VSDraw1.ScaleHeight = -8000 '-19530
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        MsgBox "このウィンドウは、「自動計測」ウィンドウを「終了」すると閉じます。", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub mnuSet_Click()
    frmKeijiSet.Show
End Sub
