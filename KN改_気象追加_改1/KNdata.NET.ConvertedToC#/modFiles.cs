using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace KNdata
{
	static class modFiles
	{

		public static string FTPpathname(string tFilename, ref string sYY, ref string sMM, ref string sDD)
		{
			//ファイル名から目的のFTPディレクトリ名を生成

			//    Dim sYY As String
			//    Dim sMM As String
			//    Dim sDD As String
			string sNN = null;

			//2009-10-12_10-00.dat
			sYY = tFilename.Substring(0, 4);
			sMM = tFilename.Substring( 5, 2);
			sDD = tFilename.Substring(8, 2);
			sNN = "/" + sYY + "/" + sMM + "/" + sDD;

			return sNN;

		}

		public static void s_ShellSort(ref string[] sArray, int Num)
		{
			int Span = 0;
			int i = 0;
			int j = 0;
			string TMP = null;

			Span = Num / 2;
			while (Span > 0) {
				for (i = Span; i <= Num - 1; i++) {
					j = i - Span + 1;
					for (j = (i - Span + 1); j >= 0; j += -Span) {
						if (sArray[j].CompareTo(sArray[j + Span]) < 1)
							break; // TODO: might not be correct. Was : Exit For
						// 順番の異なる配列要素を入れ替えます.
						TMP = sArray[j];
						sArray[j] = sArray[j + Span];
						sArray[j + Span] = TMP;
					}
				}
				Span = Span / 2;
			}
		}
//
// ファイル名を取り出す。
//
		public static string FindFileName(string strFileName)
		{
			//ファイル名の取得
			return System.IO.Path.GetFileName(strFileName);
		}

//
// パスだけを取り出す。
//
		public static string RemoveFileSpec(string strPath)
		{
			// strPath : フルパスのファイル名
			// 戻り値  : パス名
			return System.IO.Path.GetDirectoryName(strPath);
		}


		public static void sFileDelete(ref string DelFile)
		{
			KNdata.My.MyProject.Computer.FileSystem.DeleteFile(DelFile, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
		}

		public static void sFileMove(ref string DelFile)
		{
			System.IO.File.Move(DelFile, modMain.cuDir + "\\tmp\\" + FindFileName(DelFile));
		}

	}
}
