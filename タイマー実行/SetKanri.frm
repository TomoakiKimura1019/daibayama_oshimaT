VERSION 5.00
Object = "{13E51000-A52B-11D0-86DA-00608CB9FBFB}#5.0#0"; "vcf15.ocx"
Begin VB.Form frmSetKanri 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�Ǘ��l�ݒ�"
   ClientHeight    =   2880
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6150
   Icon            =   "SetKanri.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6150
   StartUpPosition =   2  '��ʂ̒���
   Begin VCF150Ctl.F1Book F1Book1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
      _0              =   $"SetKanri.frx":0442
      _1              =   $"SetKanri.frx":084B
      _2              =   $"SetKanri.frx":0C54
      _3              =   $"SetKanri.frx":105D
      _4              =   $"SetKanri.frx":1466
      _5              =   $"SetKanri.frx":1870
      _count          =   6
      _ver            =   2
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   20000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   20000
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnuFile 
      Caption         =   "̧��"
      Begin VB.Menu mnuSave 
         Caption         =   "�ۑ�"
      End
      Begin VB.Menu brank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "���"
      End
      Begin VB.Menu blank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "�I��"
      End
   End
End
Attribute VB_Name = "frmSetKanri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public OWARI As Boolean

Private Sub FileSave(SW As Boolean)
    Dim i As Integer, f As Integer, co As Integer, L As String
    Dim DM(2, 50) As String
    Dim dan_ID As Integer, kou_ID As Integer
    Dim sNO As Integer
    
    Dim Incel As Integer
    
    sNO = F1Book1.Sheet
    SW = False
    
    F1Book1.Sheet = 1
    For i = 1 To F1Book1.MaxRow
        If F1Book1.TextRC(i, 3) = "" Then SW = True: Incel = 3: Exit For
        If F1Book1.TextRC(i, 4) = "" Then SW = True: Incel = 4: Exit For
    Next i
    If SW = True Then
        F1Book1.SetActiveCell i, Incel
        MsgBox "�󔒂̃Z����������܂����B�K�����l����͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
        Exit Sub
    End If
    F1Book1.Sheet = 2
    For i = 1 To F1Book1.MaxRow
        If F1Book1.TextRC(i, 2) = "" Then SW = True: Incel = 2: Exit For
        If F1Book1.TextRC(i, 3) = "" Then SW = True: Incel = 3: Exit For
    Next i
    If SW = True Then
        F1Book1.SetActiveCell i, Incel
        MsgBox "�󔒂̃Z����������܂����B�K�����l����͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
        Exit Sub
    End If
    
    
    
    i = FileCheck(KEISOKU.Tabl_path & "�Ǘ��l�L�k.dat", "�Ǘ��l�f�[�^")
    If i = 0 Then Unload �v��Form: End
    i = FileCheck(KEISOKU.Tabl_path & "�Ǘ��l�p�C�v�c��.dat", "�Ǘ��l�f�[�^")
    If i = 0 Then Unload �v��Form: End
    
    
    kou_ID = 1
    co = 0
    f = FreeFile
    Open KEISOKU.Tabl_path & "�Ǘ��l�L�k.dat" For Input Shared As #f
        Do While Not (EOF(f))
            Line Input #f, L
            If Left$(L, 1) = ":" Then
                co = co + 1
                DM(kou_ID, co) = L
            End If
        Loop
    Close #f
    Open KEISOKU.Tabl_path & "�Ǘ��l�L�k.dat" For Output Lock Write As #f
        For i = 1 To co
            Print #f, DM(kou_ID, i)
        Next i
    
        F1Book1.Sheet = kou_ID
        For dan_ID = 1 To DanSet(kou_ID, 0).dan
            For i = 1 To 2
                Print #f, Format$(dan_ID, "@@@@");
                Print #f, Format$(i, "@@@@");
                If F1Book1.TextRC((dan_ID - 1) * 2 + i, 3) = "" Or F1Book1.TextRC((dan_ID - 1) * 2 + i, 4) = "" Then
                    Print #f, " 999  999999"
                Else
                    Print #f, Format$(F1Book1.TextRC((dan_ID - 1) * 2 + i, 3), "@@@@");
                    Print #f, Format$(F1Book1.TextRC((dan_ID - 1) * 2 + i, 4), "@@@@@@@@")
                End If
            Next i
        Next dan_ID
    Close #f
    
    kou_ID = 2
    co = 0
    f = FreeFile
    Open KEISOKU.Tabl_path & "�Ǘ��l�p�C�v�c��.dat" For Input Shared As #f
        Do While Not (EOF(f))
            Line Input #f, L
            If Left$(L, 1) = ":" Then
                co = co + 1
                DM(kou_ID, co) = L
            End If
        Loop
    Close #f
    Open KEISOKU.Tabl_path & "�Ǘ��l�p�C�v�c��.dat" For Output Lock Write As #f
        For i = 1 To co
            Print #f, DM(kou_ID, i)
        Next i
    
        F1Book1.Sheet = kou_ID
        For dan_ID = 1 To DanSet(kou_ID, 0).dan
            Print #f, Format$(dan_ID, "@@@@");
            If F1Book1.TextRC(dan_ID, 2) = "" Or F1Book1.TextRC(dan_ID, 3) = "" Then
                Print #f, "  999999  999999"
            Else
                Print #f, Format$(F1Book1.TextRC(dan_ID, 2), "@@@@@@@@");
                Print #f, Format$(F1Book1.TextRC(dan_ID, 3), "@@@@@@@@")
            End If
        Next dan_ID
    Close #f
    
    F1Book1.Modified = False
    F1Book1.Sheet = sNO
End Sub


Private Sub F1Book1_SafeEndEdit(EditString As VCF150Ctl.IF1EventArg, CancelFlag As VCF150Ctl.IF1EventArg)
On Error Resume Next
    If IsNumeric(EditString) = False Then
        MsgBox "���l�ƔF���ł��Ȃ��l�����͂���܂����B������x�A���͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
        F1Book1.CancelEdit
    Else
        If CSng(EditString) = 0 Then
            MsgBox "0�ȊO�̐��l����͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
            F1Book1.CancelEdit
        End If
    End If
End Sub

'Private Sub F1Book1_SelChange()
'    Select Case F1Book1.Col
'    Case 5
'        Label1.Caption = "�u�x�񽲯��v�́A�Z�����_�u���N���b�N����ƕύX���܂��B"
'    Case Else
'        Label1.Caption = "�Ǘ��l����͂��܂��B�l�́A�����ʁiSI�P�ʁj�œ��͂��܂��B"
'    End Select
'End Sub

Private Sub Form_Load()
    Dim dan_ID As Integer, kou_ID As Integer
    
    frmCLOSE.setKanri = False
    Width = 6240 '10425
    Height = 3540
    
    With F1Book1
        kou_ID = 1
        .Sheet = kou_ID
        .SheetName(kou_ID) = Trim$(kou(kou_ID, 1).ti1)
        For dan_ID = 1 To DanSet(kou_ID, 0).dan
            .TextRC((dan_ID - 1) * 2 + 1, 1) = Trim$(DanSet(kou_ID, dan_ID).ti)
            .TextRC((dan_ID - 1) * 2 + 2, 1) = Trim$(DanSet(kou_ID, dan_ID).ti)
            .TextRC((dan_ID - 1) * 2 + 1, 2) = "�P���Ǘ��l"
            .TextRC((dan_ID - 1) * 2 + 2, 2) = "�Q���Ǘ��l"
            If Kanri(kou_ID, dan_ID).Lebel(1) = 0 Then
                .TextRC((dan_ID - 1) * 2 + 1, 3) = 999
                .TextRC((dan_ID - 1) * 2 + 2, 3) = 999
                .TextRC((dan_ID - 1) * 2 + 1, 4) = 999999
                .TextRC((dan_ID - 1) * 2 + 2, 4) = 999999
            Else
                .TextRC((dan_ID - 1) * 2 + 1, 3) = Kanri(kou_ID, dan_ID).Hday(1)
                .TextRC((dan_ID - 1) * 2 + 2, 3) = Kanri(kou_ID, dan_ID).Hday(2)
                .TextRC((dan_ID - 1) * 2 + 1, 4) = Kanri(kou_ID, dan_ID).Lebel(1)
                .TextRC((dan_ID - 1) * 2 + 2, 4) = Kanri(kou_ID, dan_ID).Lebel(2)
            End If
        Next dan_ID
        .MaxRow = DanSet(kou_ID, 0).dan * 2
        .EnableProtection = True
        .DoSafeEvents = True
        .SetActiveCell 1, 3
        
        kou_ID = 2
        .Sheet = kou_ID
        .SheetName(kou_ID) = Trim$(kou(kou_ID, 1).ti1)
        For dan_ID = 1 To DanSet(kou_ID, 0).dan
            .TextRC(dan_ID, 1) = Trim$(DanSet(kou_ID, dan_ID).ti)
            If Kanri(kou_ID, dan_ID).Lebel(1) = 0 Then
                .TextRC(dan_ID, 2) = 999999
            Else
                .TextRC(dan_ID, 2) = Kanri(kou_ID, dan_ID).Lebel(1)
            End If
            If Kanri(kou_ID, dan_ID).Lebel(2) = 0 Then
                .TextRC(dan_ID, 3) = 999999
            Else
                .TextRC(dan_ID, 3) = Kanri(kou_ID, dan_ID).Lebel(2)
            End If
        Next dan_ID
        .MaxRow = DanSet(kou_ID, 0).dan
        .EnableProtection = True
        .DoSafeEvents = True
        .SetActiveCell 1, 2
    End With
    
    F1Book1.Sheet = 1
    F1Book1.Modified = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Response As Integer
    Dim SW As Boolean
    If F1Book1.Modified Then
        Response = MsgBox("�ύX���ۑ�����Ă��܂���B�ۑ����܂����H", vbYesNoCancel + vbExclamation, "�I���̊m�F")
        If Response = vbCancel Then Cancel = True: Exit Sub
        If Response = vbYes Then
            Call FileSave(SW)
            If SW = True Then Cancel = True: Exit Sub
        End If
    End If
    frmCLOSE.setKanri = True
End Sub

Private Sub mnuEnd_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
On Error GoTo trap1
    F1Book1.SetSelection 1, 1, F1Book1.MaxRow, F1Book1.MaxCol
    F1Book1.SetPrintAreaFromSelection
    F1Book1.FilePrintEx True, False
    If F1Book1.Sheet = 1 Then F1Book1.SetActiveCell 1, 3
    If F1Book1.Sheet = 1 Then F1Book1.SetActiveCell 1, 2
    Me.Refresh
trap1:
End Sub

Private Sub mnuSave_Click()
    Dim SW As Boolean
    
    If F1Book1.Modified = True Then Call FileSave(SW)
    If SW = False Then MsgBox "�ۑ����I�����܂����B", vbInformation
End Sub

