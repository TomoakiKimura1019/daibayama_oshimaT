Option Strict Off
Option Explicit On
Module WinAPI
	
	'�l�b�g���[�N�h���C�u��ChDrive����
	Public Declare Function SetCurrentDirectory Lib "kernel32"  Alias "SetCurrentDirectoryA"(ByVal CurrentDir As String) As Integer
	
	'INI�t�@�C�����ǂݍ���
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'INI�t�@�C���ɏ�������
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	'UPGRADE_ISSUE: �p�����[�^ 'As Any' �̐錾�̓T�|�[�g����܂���B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' ���N���b�N���Ă��������B
	Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	
	Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Public Declare Function GetForegroundWindow Lib "user32" () As Integer
	Public Declare Function GetActiveWindow Lib "user32" () As Integer
	Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Integer) As Integer
	'UPGRADE_WARNING: �\���� RECT �ɁA���� Declare �X�e�[�g�����g�̈����Ƃ��ă}�[�V�������O������n���K�v������܂��B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"' ���N���b�N���Ă��������B
	Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hdc As Integer) As Integer
	
	Public Structure RECT
		'UPGRADE_NOTE: Left �� Left_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right �� Right_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Public Const SRCCOPY As Integer = &HCC0020 ' (DWORD) dest = source
	
	Public Const SW_SHOW As Short = 5
	
	'�f�X�N�g�b�v�E�B���h�E�̃n���h�����擾����API
	Public Declare Function GetDesktopWindow Lib "user32" () As Integer
	
	'�w�肳�ꂽ�t�@�C�����I�[�v���A���邢�͕\������API
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	'
	'�� ���ӁI�I    �����i�[����o�b�t�@�͕K���������w�肷�邱�ƁB
	'               WritePrivateProfileString�Ńt�@�C�����Ȃ������ꍇ�͍쐬�����B
	'
	'<<���>>GetPrivateProfileString( �Z�N�V���� as String
	'                                 �L�[ as String
	'                                 �f�t�H���g as String          �L�[���Ȃ������ꍇ�̒l
	'                                 ��� as Any                   �����i�[����o�b�t�@
	'                                 ���̒��� as Long            �o�b�t�@�̒���
	'                                 INI�t�@�C���p�X as String
	'
	'<<���>>WritePrivateProfileString( �Z�N�V���� as String
	'                                   �L�[ as Any
	'                                   �������ޏ�� as String
	'                                   INI�t�@�C���p�X as String
	'
	
	'**********************************
	'        INI�t�@�C���ǂݍ���
	'**********************************
	Public Function GetIni(ByRef section As String, ByRef key As String, ByRef INIFile As String) As String
		Dim StrBuf As New VB6.FixedLengthString(1024)
		Dim ret As Integer
		Dim EBuf As String
		Dim i As Integer
		
		StrBuf.Value = New String(Chr(0), 1024)
		ret = GetPrivateProfileString(section, key, "", StrBuf.Value, 1024, INIFile)
		
		i = InStr(StrBuf.Value, vbNullChar)
		If i <> 0 Then
			EBuf = Left(StrBuf.Value, i - 1)
		Else
			EBuf = StrBuf.Value
		End If
		
		GetIni = EBuf
		
	End Function
	
	'**********************************
	'        INI�t�@�C����������
	'**********************************
	'UPGRADE_NOTE: str �� str_Renamed �ɃA�b�v�O���[�h����܂����B �ڍׂɂ��ẮA'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' ���N���b�N���Ă��������B
	Public Function WriteIni(ByRef section As String, ByRef key As String, ByRef str_Renamed As String, ByRef INIFile As String) As Integer
		Dim ret As Integer
		ret = WritePrivateProfileString(section, key, str_Renamed, INIFile)
		WriteIni = ret
	End Function
End Module