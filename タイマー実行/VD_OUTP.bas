Attribute VB_Name = "VD_OUTP"
Option Explicit

Public Const ANGINC As Double = PI / 180#

Public Sub VD_AnkCsize(TARGETOBJECT As Object, ByVal PANKSIZE As Integer)
    TARGETOBJECT.FontSize = PANKSIZE
End Sub

Public Sub VD_CIL(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single, _
                    ByVal RG As Single, ByVal CSA As Integer, ByVal CEA As Integer)
    
    If CSA = 0 And CEA = 360 Then
        TARGETOBJECT.DrawCircle PX, PY, RG
    Else
        TARGETOBJECT.DrawCircle PX, PY, RG, CSA * ANGINC, CEA * ANGINC
    End If
End Sub

Public Sub VD_DDD(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single)
    TARGETOBJECT.DrawLine PX, PY
End Sub

Public Sub VD_KPUT(TARGETOBJECT As Object, ByVal SS As String, _
            ByVal PX As Single, ByVal PY As Single, ByVal SIZ As Integer, ByVal XOFF As Integer, ByVal YOFF As Integer, Ang As Integer, BKstyle As Integer)
    
    Dim cr As String
    Dim i As Integer
    
'    If frmSakuzu.PRINTOUT = 1 Then PY = PY + 5
'    If frmSakuzu.PRINTOUT = 1 Then
'        fntText.Size = SIZ - 1
'    Else
'        fntText.Size = SIZ
'    End If
'
'    'fntText.Size = SIZ
'    For i = 1 To Len(SS)
'        cr = Mid$(SS, i, 1)
'        TARGETOBJECT.CurrentX = PX
'        TARGETOBJECT.CurrentY = PY
'        TARGETOBJECT.Print cr
'        If 0 < Asc(cr$) And Asc(cr$) <= 255 Then
'            PX = PX + (XOFF / 2)
'        Else
'            PX = PX + XOFF
'        End If
'        'PX = PX + XOFF '* 6
'        PY = PY + YOFF '* 6
'    Next i
    
    TARGETOBJECT.FontSize = SIZ
    TARGETOBJECT.TextAngle = Ang * 10
    TARGETOBJECT.BackStyle = BKstyle

    
    For i = 1 To Len(SS)
        cr = Mid$(SS, i, 1)
        
        TARGETOBJECT.X1 = PX
        TARGETOBJECT.Y1 = PY
        TARGETOBJECT.TextOut = cr
        If 0 < Asc(cr$) And Asc(cr$) <= 255 Then
            PX = PX + (XOFF / 2)
        Else
            PX = PX + XOFF
        End If
        'PX = PX + XOFF '* 6
        PY = PY + YOFF '* 6
    Next i
    

End Sub

Public Sub VD_LTCD(TARGETOBJECT As Object, ByVal SENC As Integer)
    Dim SENSYU(10) As Long
    SENSYU(0) = 0
    SENSYU(1) = 1
    SENSYU(2) = 2
    SENSYU(3) = 3
    SENSYU(4) = 4
    TARGETOBJECT.PenStyle = SENSYU(SENC)
End Sub

Public Sub VD_MK(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single, ByVal MKF As Integer)
    Dim R As Single, MX As Single, MY As Single
    Dim MKFF As Integer
    
    R = 8 * 5
    R = 12 * 5
    MX = PX: MY = PY
    
    MKFF = MKF Mod 7
    If MKF <> 0 And MKFF = 0 Then MKFF = 7
    
    Select Case MKFF
    Case 1
        Call VD_CIL(TARGETOBJECT, PX, PY, R, 0, 360)
    Case 2
        TARGETOBJECT.DrawLine MX - R, MY + R, MX + R, MY - R
        TARGETOBJECT.DrawLine MX + R, MY + R, MX - R, MY - R
    Case 3
        TARGETOBJECT.DrawRectangle MX - R, MY + R, MX + R, MY - R
    Case 4
        PX = MX: PY = MY + R: Call VD_MMM(TARGETOBJECT, PX, PY)
        PX = MX - Sqr(3) / 2 * R: PY = MY - R / 2: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX + Sqr(3) / 2 * R: PY = MY - R / 2: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY + R: Call VD_DDD(TARGETOBJECT, PX, PY)
    Case 5
        PX = MX - R: PY = MY + R: Call VD_MMM(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY + R: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY - R: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY - R: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY + R: Call VD_DDD(TARGETOBJECT, PX, PY)
    Case 6
        PX = MX: PY = MY - R: Call VD_MMM(TARGETOBJECT, PX, PY)
        PX = MX - Sqr(3) / 2 * R: PY = MY + R / 2: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX + Sqr(3) / 2 * R: PY = MY + R / 2: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY - R: Call VD_DDD(TARGETOBJECT, PX, PY)
    Case 7
        PX = MX: PY = MY + R: Call VD_MMM(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY - R: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY: Call VD_DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY + R: Call VD_DDD(TARGETOBJECT, PX, PY)
    End Select
        
    Call VD_MMM(TARGETOBJECT, MX, MY)
End Sub

Public Sub VD_MMM(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single)
    TARGETOBJECT.DrawLine PX, PY, PX, PY
End Sub

Public Sub VD_PPANK(TARGETOBJECT As Object, ByVal PANK As String, ByVal PX As Single, ByVal PY As Single)
    
'    If frmSakuzu.PRINTOUT = 1 Then PY = PY + 5
    
'    TARGETOBJECT.CurrentX = PX
'    TARGETOBJECT.CurrentY = PY
'    TARGETOBJECT.Print PANK
    
    TARGETOBJECT.TextAngle = 0
    TARGETOBJECT.X1 = PX
    TARGETOBJECT.Y1 = PY
    TARGETOBJECT.TextOut = PANK
End Sub

Public Sub VD_PANKFV(TARGETOBJECT As Object, ByVal PANKFM As String, ByVal PANKF As Variant, ByVal PX As Single, ByVal PY As Single)
    
    Dim MyStr As String
    Dim Xkankaku As Integer
    
    'If frmSakuzu.PRINTOUT = 1 Then PY = PY + 5: Xkankaku = 20 Else Xkankaku = 15
    
    MyStr = Format$(PANKF, PANKFM)
    MyStr = Right$(Space$(Len(PANKFM)) + MyStr, Len(PANKFM))
    PX = PX - Len(MyStr) * Xkankaku
'    TARGETOBJECT.CurrentX = PX
'    TARGETOBJECT.CurrentY = PY
'    TARGETOBJECT.Print MyStr
    
    TARGETOBJECT.TextAngle = 0
    TARGETOBJECT.X1 = PX
    TARGETOBJECT.Y1 = PY
    TARGETOBJECT.TextOut = MyStr
End Sub

Public Sub VD_PENJ(TARGETOBJECT As Object, ByVal PENC As Integer)
    Dim IRO(10) As Long
    Dim PENCC As Integer
    
    PENCC = PENC Mod 9
    If PENC <> 0 And PENCC = 0 Then PENCC = 9
    IRO(1) = RGB(0, 0, 0)         '黒
    IRO(2) = RGB(256, 0, 0)       '赤
    IRO(3) = RGB(0, 0, 128)       '青   RGB(0, 0, 256)       '濃紺
    IRO(4) = RGB(0, 256, 0)       '黄緑 RGB(0, 128, 0)       '緑
    IRO(5) = RGB(0, 256, 256)     '水色
    IRO(6) = RGB(256, 0, 256)     'ピンク
    IRO(7) = RGB(256, 240, 0)     '黄色 RGB(256, 160, 0)
    IRO(8) = RGB(128, 0, 256)     '紫
    IRO(9) = RGB(128, 64, 64)     '茶色
    
    TARGETOBJECT.PenColor = IRO(PENCC)
    TARGETOBJECT.TextColor = IRO(PENCC)
End Sub

Public Sub VD_BRectangle(TARGETOBJECT As Object, ByVal PX1 As Single, ByVal PY1 As Single, ByVal PX2 As Single, ByVal PY2 As Single, Bstyle As Integer)
    Dim PPX1 As Single
    Dim PPY1 As Single
    Dim PPX2 As Single
    Dim PPY2 As Single

    If Bstyle = 0 Then TARGETOBJECT.BrushStyle = bsSolid
    If Bstyle = 1 Then TARGETOBJECT.BrushStyle = bsDiagonalUp
    
    PPX1 = PX1 * 5.67
    PPY1 = TARGETOBJECT.PageHeight - PY1 * 5.67
    PPX2 = PX2 * 5.67
    PPY2 = TARGETOBJECT.PageHeight - PY2 * 5.67
'    TARGETOBJECT.CurrentX = PX1
'    TARGETOBJECT.CurrentY = PY1
    TARGETOBJECT.DrawRectangle PX1, PY1, PX2, PY2
    TARGETOBJECT.BrushStyle = bsSolid
End Sub


