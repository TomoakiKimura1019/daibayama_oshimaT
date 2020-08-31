VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPoSet0 
   Caption         =   "à íuê›íË"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command2 
      Caption         =   "ï€ë∂"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Index           =   0
      Left            =   5520
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "∑¨›æŸ"
      Height          =   495
      Index           =   1
      Left            =   5520
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ì«Ç›çûÇ›"
         Height          =   495
         Left            =   3960
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorBkg    =   12632256
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "ì«Ç›çûÇ›à íu"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   525
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPoSet0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Posave()
    Dim f As Integer, i As Integer
    Dim SS1 As String, SS2 As String
    
    f = FreeFile
    Open InitDT.PoFILE For Output Lock Write As #f
    For i = 1 To InitDT.CO
        If MSFlexGrid1.TextMatrix(i, 1) = "" Or MSFlexGrid1.TextMatrix(i, 1) = "******" Then
            PoDT(i).Hdt = -999
        Else
            PoDT(i).Hdt = MSFlexGrid1.TextMatrix(i, 1)
        End If
        If MSFlexGrid1.TextMatrix(i, 2) = "" Or MSFlexGrid1.TextMatrix(i, 2) = "******" Then
            PoDT(i).Vdt = -999
        Else
            PoDT(i).Vdt = MSFlexGrid1.TextMatrix(i, 2)
        End If
        If MSFlexGrid1.TextMatrix(i, 3) = "" Or MSFlexGrid1.TextMatrix(i, 3) = "******" Then
            PoDT(i).Sdt = -999
        Else
            PoDT(i).Sdt = MSFlexGrid1.TextMatrix(i, 3)
        End If
        
        SS1 = Format(i, "@@@@")
        SS2 = CStr(PoDT(i).Hdt)
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        SS2 = CStr(PoDT(i).Vdt)
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        SS2 = CStr(PoDT(i).Sdt)
        SS1 = SS1 & Space$(12 - LenB(StrConv(SS2, vbFromUnicode))) & SS2
        Print #f, SS1
    Next i
    Close #f
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then Call Posave
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
    Call Posave
End Sub

Private Sub Command5_Click()
    Dim no As Integer
    
    no = Combo1.ListIndex + 1
    
    MSFlexGrid1.TextMatrix(no, 1) = 1.2345678
    MSFlexGrid1.TextMatrix(no, 2) = 2.2345678
    MSFlexGrid1.TextMatrix(no, 3) = 3.2345678
End Sub

Private Sub Form_Load()
    Dim f As Integer, L As String, i As Integer
        
    Top = frmSyokiset.Top + frmSyokiset.Height - 1000
    Left = frmSyokiset.Left + frmSyokiset.Width - 1000
    
    For i = 1 To InitDT.CO
        Combo1.AddItem "No." & CStr(i)
    Next i
    Combo1.ListIndex = 0
    
    MSFlexGrid1.Rows = InitDT.CO + 1
    MSFlexGrid1.ColWidth(0) = 600
    MSFlexGrid1.ColWidth(1) = 1400
    MSFlexGrid1.ColWidth(2) = 1400
    MSFlexGrid1.ColWidth(3) = 1400
    MSFlexGrid1.ColAlignment(0) = 4
    For i = 1 To MSFlexGrid1.Cols - 1
        MSFlexGrid1.ColAlignment(i) = 7
        MSFlexGrid1.Row = 0: MSFlexGrid1.Col = i: MSFlexGrid1.CellAlignment = 4
    Next i
    MSFlexGrid1.TextMatrix(0, 1) = "H (ïb)"
    MSFlexGrid1.TextMatrix(0, 2) = "V (ïb)"
    MSFlexGrid1.TextMatrix(0, 3) = "S (Çç)"
    
    Erase PoDT
    
    For i = 1 To InitDT.CO
        MSFlexGrid1.TextMatrix(i, 0) = "No." & CStr(i)
        PoDT(i).Hdt = -999
        PoDT(i).Vdt = -999
        PoDT(i).Sdt = -999
    Next i
    
    If Dir(InitDT.PoFILE) <> "" Then
        i = 0
        f = FreeFile
        Open InitDT.PoFILE For Input Shared As #f
        Do While Not (EOF(f))
            Line Input #f, L
            i = i + 1
            PoDT(i).Hdt = CDbl(Mid(L, 5, 12))
            PoDT(i).Vdt = CDbl(Mid(L, 17, 12))
            PoDT(i).Sdt = CDbl(Mid(L, 29, 12))
            
            If PoDT(i).Hdt = -999 Then
                MSFlexGrid1.TextMatrix(i, 1) = "******"
            Else
                MSFlexGrid1.TextMatrix(i, 1) = PoDT(i).Hdt
            End If
            If PoDT(i).Vdt = -999 Then
                MSFlexGrid1.TextMatrix(i, 2) = "******"
            Else
                MSFlexGrid1.TextMatrix(i, 2) = PoDT(i).Vdt
            End If
            If PoDT(i).Sdt = -999 Then
                MSFlexGrid1.TextMatrix(i, 3) = "******"
            Else
                MSFlexGrid1.TextMatrix(i, 3) = PoDT(i).Sdt
            End If
        Loop
        Close #f
    End If

End Sub

