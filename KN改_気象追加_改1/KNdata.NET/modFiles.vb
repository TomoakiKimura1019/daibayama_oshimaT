Option Strict Off
Option Explicit On
Module modFiles
	
    Public Function FTPpathname(ByVal tFilename As String, ByRef sYY As String, ByRef sMM As String, ByRef sDD As String) As String
        '�t�@�C��������ړI��FTP�f�B���N�g�����𐶐�

        '    Dim sYY As String
        '    Dim sMM As String
        '    Dim sDD As String
        Dim sNN As String

        '2009-10-12_10-00.dat
        sYY = Mid(tFilename, 1, 4)
        sMM = Mid(tFilename, 6, 2)
        sDD = Mid(tFilename, 9, 2)
        sNN = "/" & sYY & "/" & sMM & "/" & sDD

        FTPpathname = sNN

    End Function

    Public Sub s_ShellSort(ByRef sArray() As String, ByVal Num As Integer)
        Dim Span As Integer
        Dim i As Integer
        Dim j As Integer
        Dim TMP As String

        Span = Num \ 2
        Do While Span > 0
            For i = Span To Num - 1
                j = i - Span + 1
                For j = (i - Span + 1) To 0 Step -Span
                    If sArray(j) <= sArray(j + Span) Then Exit For
                    ' ���Ԃ̈قȂ�z��v�f�����ւ��܂�.
                    TMP = sArray(j)
                    sArray(j) = sArray(j + Span)
                    sArray(j + Span) = TMP
                Next j
            Next i
            Span = Span \ 2
        Loop
    End Sub
	'
	' �t�@�C���������o���B
	'
	Public Function FindFileName(ByVal strFileName As String) As String
        '�t�@�C�����̎擾
        FindFileName = System.IO.Path.GetFileName(strFileName)
    End Function
	
	'
	' �p�X���������o���B
	'
	Public Function RemoveFileSpec(ByVal strPath As String) As String
		' strPath : �t���p�X�̃t�@�C����
		' �߂�l  : �p�X��

        RemoveFileSpec = System.IO.Path.GetDirectoryName(strPath)

    End Function
	
	Public Sub sFileDelete(ByRef DelFile As String)

        My.Computer.FileSystem.DeleteFile( _
      DelFile, _
      FileIO.UIOption.OnlyErrorDialogs, _
      FileIO.RecycleOption.SendToRecycleBin)

    End Sub
	
	Public Sub sFileMove(ByRef DelFile As String)
        System.IO.File.Move(DelFile, cuDir & "\tmp\" & FindFileName(DelFile))
    End Sub
	
End Module