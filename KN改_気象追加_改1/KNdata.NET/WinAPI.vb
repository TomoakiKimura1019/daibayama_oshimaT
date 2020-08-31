Option Strict Off
Option Explicit On
Module WinAPI
	
	'ネットワークドライブにChDriveする
	Public Declare Function SetCurrentDirectory Lib "kernel32"  Alias "SetCurrentDirectoryA"(ByVal CurrentDir As String) As Integer
	
	'INIファイルより読み込み
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'INIファイルに書き込み
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	'UPGRADE_ISSUE: パラメータ 'As Any' の宣言はサポートされません。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"' をクリックしてください。
	Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	
	Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
	Public Declare Function GetForegroundWindow Lib "user32" () As Integer
	Public Declare Function GetActiveWindow Lib "user32" () As Integer
	Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Integer) As Integer
	'UPGRADE_WARNING: 構造体 RECT に、この Declare ステートメントの引数としてマーシャリング属性を渡す必要があります。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"' をクリックしてください。
	Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hdc As Integer) As Integer
	
	Public Structure RECT
		'UPGRADE_NOTE: Left は Left_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right は Right_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Public Const SRCCOPY As Integer = &HCC0020 ' (DWORD) dest = source
	
	Public Const SW_SHOW As Short = 5
	
	'デスクトップウィンドウのハンドルを取得するAPI
	Public Declare Function GetDesktopWindow Lib "user32" () As Integer
	
	'指定されたファイルをオープン、あるいは表示するAPI
	Public Declare Function ShellExecute Lib "shell32.dll"  Alias "ShellExecuteA"(ByVal hWnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
	'
	'※ 注意！！    情報を格納するバッファは必ず長さを指定すること。
	'               WritePrivateProfileStringでファイルがなかった場合は作成される。
	'
	'<<解説>>GetPrivateProfileString( セクション as String
	'                                 キー as String
	'                                 デフォルト as String          キーがなかった場合の値
	'                                 情報 as Any                   情報を格納するバッファ
	'                                 情報の長さ as Long            バッファの長さ
	'                                 INIファイルパス as String
	'
	'<<解説>>WritePrivateProfileString( セクション as String
	'                                   キー as Any
	'                                   書き込む情報 as String
	'                                   INIファイルパス as String
	'
	
	'**********************************
	'        INIファイル読み込み
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
	'        INIファイル書き込み
	'**********************************
	'UPGRADE_NOTE: str は str_Renamed にアップグレードされました。 詳細については、'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"' をクリックしてください。
	Public Function WriteIni(ByRef section As String, ByRef key As String, ByRef str_Renamed As String, ByRef INIFile As String) As Integer
		Dim ret As Integer
		ret = WritePrivateProfileString(section, key, str_Renamed, INIFile)
		WriteIni = ret
	End Function
End Module