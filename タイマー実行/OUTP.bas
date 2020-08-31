Attribute VB_Name = "OUTP"
Option Explicit
'----------------------------------------------------------------------------------------------
'   作図・作表用描画コントロール
'----------------------------------------------------------------------------------------------

Public Const ANGINC As Double = PI / 180#  '度→ラジアン

'**********************************************************************************************
'   フォントサイズ
'       入力 PANKSIZE：フォントサイズ
'**********************************************************************************************
Public Sub AnkCsize(TARGETOBJECT As Object, ByVal PANKSIZE As Integer)
    TARGETOBJECT.FontSize = PANKSIZE
End Sub

'**********************************************************************************************
'   円、弧、扇形を描画
'       入力 PX:Ｘ座標
'            PY:Ｙ座標
'            RG:角度
'            CSA:楕円の開始点
'            CEA:楕円の終了点
'**********************************************************************************************
Public Sub CIL(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single, _
                    ByVal RG As Single, ByVal CSA As Integer, ByVal CEA As Integer)
    Dim PPX As Single
    Dim PPY As Single
    Dim RRG As Long
    
    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67
    RRG = RG * 5.67
    
    If CSA = 0 And CEA = 360 Then
        TARGETOBJECT.DrawCircle PPX, PPY, RRG
    Else
        TARGETOBJECT.DrawCircle PPX, PPY, RRG, CSA * ANGINC, CEA * ANGINC
    End If
End Sub

'**********************************************************************************************
'   線を描画
'       入力  PX:Ｘ座標
'             PY:Ｙ座標
'**********************************************************************************************
Public Sub DDD(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single)
    Dim PPX As Single
    Dim PPY As Single

    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67
    
    TARGETOBJECT.DrawLine PPX, PPY

End Sub

'**********************************************************************************************
'   文字を描画（文字の間隔などを設定可能）
'       入力  SS:描画文字
'             PX:Ｘ座標
'             PY:Ｙ座標
'             SIZ:フォントサイズ
'             XOFF:Ｘ方向文字間隔
'             YOFF:Ｙ方向文字間隔
'             ANG:文字の回転（-3,600～3,600：0.1度単位）
'**********************************************************************************************
Public Sub KPUT(TARGETOBJECT As Object, ByVal SS As String, _
            ByVal PX As Single, ByVal PY As Single, ByVal SIZ As Integer, ByVal XOFF As Integer, ByVal YOFF As Integer, Ang As Integer)
    
    Dim cr As String
    Dim i As Integer, j As Integer
    Dim PPX As Single
    Dim PPY As Single
    Dim PXOFF As Single
    Dim PYOFF As Single

    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67
    PXOFF = XOFF * 5.67: PYOFF = YOFF * (-5.67)
    
    TARGETOBJECT.FontSize = SIZ
    TARGETOBJECT.TextAngle = Ang * 10
    For i = 1 To Len(SS)
        cr = Mid$(SS, i, 1)
        
        TARGETOBJECT.CurrentX = PPX
        TARGETOBJECT.CurrentY = PPY
        TARGETOBJECT.Text = cr
        If LenB(StrConv(cr, vbFromUnicode)) = 1 Then
            PPX = PPX + (PXOFF / 2)
        Else
            PPX = PPX + PXOFF
        End If
        PPY = PPY + PYOFF
    Next i
    
End Sub

'**********************************************************************************************
'   線種設定
'       入力  SENC:0=実線
'                  1=破線
'                  2=点線
'                  3=一点鎖線
'                  4=二点鎖線
'**********************************************************************************************
Public Sub LTCD(TARGETOBJECT As Object, ByVal SENC As Integer)
    Dim SENSYU(10) As Long
    SENSYU(0) = 0
    SENSYU(1) = 1
    SENSYU(2) = 2
    SENSYU(3) = 3
    SENSYU(4) = 4
    TARGETOBJECT.PenStyle = SENSYU(SENC)
End Sub

'**********************************************************************************************
'   マーカーの描画
'       入力 PX:Ｘ座標
'            PY:Ｙ座標
'            MKF:ｽﾀｲﾙ 1=○
'                     2=×
'                     3=□
'                     4=△
'                     5=△+▽
'                     6=▽
'                     7=ひし形
'**********************************************************************************************
Public Sub MK(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single, ByVal MKF As Integer)
    Dim R As Single, MX As Single, MY As Single
    Dim MKFF As Integer
    Dim PPX As Single
    Dim PPY As Single
    Dim RR As Long
    
    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67
    R = 8: RR = R * 5.67
    
    MX = PX: MY = PY
    
    MKFF = MKF Mod 7
    If MKF <> 0 And MKFF = 0 Then MKFF = 7
    
    Select Case MKFF
    Case 1
        Call CIL(TARGETOBJECT, PX, PY, 8, 0, 360)
    Case 2
        TARGETOBJECT.DrawLine PPX - RR, PPY + RR, PPX + RR, PPY - RR
        TARGETOBJECT.DrawLine PPX + RR, PPY + RR, PPX - RR, PPY - RR
    Case 3
        TARGETOBJECT.DrawRectangle PPX - RR, PPY + RR, PPX + RR, PPY - RR
    Case 4
        PX = MX: PY = MY + R: Call MMM(TARGETOBJECT, PX, PY)
        PX = MX - Sqr(3) / 2 * R: PY = MY - R / 2: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX + Sqr(3) / 2 * R: PY = MY - R / 2: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY + R: Call DDD(TARGETOBJECT, PX, PY)
    Case 5
        PX = MX - R: PY = MY + R: Call MMM(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY + R: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY - R: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY - R: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY + R: Call DDD(TARGETOBJECT, PX, PY)
    Case 6
        PX = MX: PY = MY - R: Call MMM(TARGETOBJECT, PX, PY)
        PX = MX - Sqr(3) / 2 * R: PY = MY + R / 2: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX + Sqr(3) / 2 * R: PY = MY + R / 2: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY - R: Call DDD(TARGETOBJECT, PX, PY)
    Case 7
        PX = MX: PY = MY + R: Call MMM(TARGETOBJECT, PX, PY)
        PX = MX - R: PY = MY: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY - R: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX + R: PY = MY: Call DDD(TARGETOBJECT, PX, PY)
        PX = MX: PY = MY + R: Call DDD(TARGETOBJECT, PX, PY)
    End Select
        
    Call MMM(TARGETOBJECT, MX, MY)
End Sub

'**********************************************************************************************
'   点を描画
'       入力  PX:Ｘ座標
'             PY:Ｙ座標
'**********************************************************************************************
Public Sub MMM(TARGETOBJECT As Object, ByVal PX As Single, ByVal PY As Single)
    Dim PPX As Single
    Dim PPY As Single

    'Twips単位へ変換
    PPX = PX * 5.67
    PPY = TARGETOBJECT.PageHeight - PY * 5.67
    
    TARGETOBJECT.CurrentX = PPX
    TARGETOBJECT.CurrentY = PPY
    TARGETOBJECT.DrawLine PPX, PPY, PPX, PPY
    
End Sub

'**********************************************************************************************
'   文字を描画（文字の間隔などを設定不可）
'       入力  PANK:描画文字
'             PX:Ｘ座標
'             PY:Ｙ座標
'**********************************************************************************************
Public Sub PPANK(TARGETOBJECT As Object, ByVal PANK, ByVal PX As Single, ByVal PY As Single)
    
    Dim PPX As Single
    Dim PPY As Single

    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67
    
    TARGETOBJECT.TextAngle = 0
    TARGETOBJECT.CurrentX = CLng(PPX)
    TARGETOBJECT.CurrentY = PPY
    TARGETOBJECT.Text = PANK
End Sub

'**********************************************************************************************
'   数値→文字変換後、文字を描画（文字の間隔などを設定不可）
'       入力  PANKFM:描画数値
'             PANKF:文字フォーマット
'             PX:Ｘ座標
'             PY:Ｙ座標
'**********************************************************************************************
Public Sub PANKFV(TARGETOBJECT As Object, ByVal PANKFM As String, ByVal PANKF As Variant, _
            ByVal PX As Single, ByVal PY As Single)
    
    Dim MyStr As String
    Dim Xkankaku As Integer
    Dim PPX As Single
    Dim PPY As Single

    
    PY = PY + 5: Xkankaku = 20
    
    MyStr = Format$(PANKF, PANKFM)
    MyStr = Right$(Space$(Len(PANKFM)) + MyStr, Len(PANKFM))
    PX = PX - Len(MyStr) * Xkankaku
    
    'Twips単位へ変換
    PPX = PX * 5.67: PPY = TARGETOBJECT.PageHeight - PY * 5.67

    TARGETOBJECT.TextAngle = 0
    TARGETOBJECT.CurrentX = PPX
    TARGETOBJECT.CurrentY = PPY
    TARGETOBJECT.Text = MyStr
End Sub

'**********************************************************************************************
'   グラフィック・テキストの描画色の設定
'       入力 PENJ：色番号
'**********************************************************************************************
Public Sub PENJ(TARGETOBJECT As Object, ByVal PENC As Integer)
    Dim IRO(10) As Long
    Dim PENCC As Integer
    
'    PENCC = PENC Mod 8
'    If PENC <> 0 And PENCC = 0 Then PENCC = 8
'    IRO(1) = RGB(0, 0, 0)          '黒
'    IRO(2) = RGB(256, 0, 0)        '赤
'    IRO(3) = RGB(0, 0, 128)        '濃紺 RGB(0, 0, 256) 青
'    IRO(4) = RGB(0, 256, 0)        '黄緑 RGB(0, 128, 0) 緑
'    IRO(5) = RGB(0, 256, 256)      '水色
'    'IRO(5) = RGB(256, 0, 256)      'ピンク
'    IRO(6) = RGB(256, 128, 0)      '黄色
'    IRO(7) = RGB(128, 0, 256)      '紫
'    IRO(8) = RGB(128, 64, 64)      '茶色
    
    PENCC = PENC Mod 9
    If PENC <> 0 And PENCC = 0 Then PENCC = 9
    IRO(1) = RGB(0, 0, 0)          '黒
    IRO(2) = RGB(256, 0, 0)        '赤
    IRO(3) = RGB(0, 0, 128)        '濃紺 RGB(0, 0, 256) 青
    IRO(4) = RGB(0, 256, 0)        '黄緑 RGB(0, 128, 0) 緑
    IRO(5) = RGB(0, 256, 256)      '水色
    IRO(6) = RGB(256, 0, 256)      'ピンク
    IRO(7) = RGB(256, 200, 64)     '黄色 RGB(256, 160, 0)
    IRO(8) = RGB(128, 0, 256)      '紫
    IRO(9) = RGB(128, 64, 64)      '茶色
    'TARGETOBJECT.ForeColor = IRO(PENCC)
    TARGETOBJECT.PenColor = IRO(PENCC)
    TARGETOBJECT.TextColor = IRO(PENCC)
End Sub

'**********************************************************************************************
'   長方形､または角のまるい長方形を描画します
'       入力 PX1:左上隅のＸ座標
'            PY1:左上隅のＹ座標
'            PX2:右下隅のＸ座標
'            PY2:右下隅のＹ座標
'**********************************************************************************************
Public Sub BRectangle(TARGETOBJECT As Object, ByVal PX1 As Single, ByVal PY1 As Single, ByVal PX2 As Single, ByVal PY2 As Single, Bstyle As Integer)
    Dim PPX1 As Single
    Dim PPY1 As Single
    Dim PPX2 As Single
    Dim PPY2 As Single

    If Bstyle = 0 Then TARGETOBJECT.BrushStyle = bsSolid
    If Bstyle = 1 Then TARGETOBJECT.BrushStyle = bsDiagonalUp
    
    'Twips単位へ変換
    PPX1 = PX1 * 5.67
    PPY1 = TARGETOBJECT.PageHeight - PY1 * 5.67
    PPX2 = PX2 * 5.67
    PPY2 = TARGETOBJECT.PageHeight - PY2 * 5.67
    
    TARGETOBJECT.CurrentX = PPX1
    TARGETOBJECT.CurrentY = PPY1
    TARGETOBJECT.DrawRectangle PPX1, PPY1, PPX2, PPY2
    TARGETOBJECT.BrushStyle = bsSolid
End Sub

