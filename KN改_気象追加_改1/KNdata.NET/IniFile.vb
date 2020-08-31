Imports System.Runtime.InteropServices
Imports System.Text

''' <summary>
''' INIファイルを読み書きするクラス
''' </summary>
Public Class IniFile
    <DllImport("kernel32.dll")> _
    Private Shared Function GetPrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedstring As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    End Function

    <DllImport("kernel32.dll")> _
    Private Shared Function WritePrivateProfileString(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpstring As String, ByVal lpFileName As String) As Integer
    End Function

    Private filePath As String

    ''' <summary>
    ''' ファイル名を指定して初期化します。
    ''' ファイルが存在しない場合は初回書き込み時に作成されます。
    ''' </summary>
    Public Sub New(ByVal filePath As String)
        Me.filePath = filePath
    End Sub

    ''' <summary>
    ''' sectionとkeyからiniファイルの設定値を取得、設定します。 
    ''' </summary>
    ''' <returns>指定したsectionとkeyの組合せが無い場合は""が返ります。</returns>
    Default Public Property Item(ByVal section As String, ByVal key As String) As String
        Get
            Dim sb As New StringBuilder(256)
            GetPrivateProfileString(section, key, String.Empty, sb, sb.Capacity, filePath)
            Return sb.ToString()
        End Get
        Set(ByVal value As String)
            WritePrivateProfileString(section, key, value, filePath)
        End Set
    End Property

    ''' <summary>
    ''' sectionとkeyからiniファイルの設定値を取得します。
    ''' 指定したsectionとkeyの組合せが無い場合はdefaultvalueで指定した値が返ります。
    ''' </summary>
    ''' <returns>
    ''' 指定したsectionとkeyの組合せが無い場合はdefaultvalueで指定した値が返ります。
    ''' </returns>
    Public Function GetValue(ByVal section As String, ByVal key As String, ByVal defaultvalue As String) As String
        Dim sb As New StringBuilder(256)
        GetPrivateProfileString(section, key, defaultvalue, sb, sb.Capacity, filePath)
        Return sb.ToString()
    End Function
End Class
