using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

 // ERROR: Not supported in C#: OptionDeclaration
using System.IO;
using System.Text;

using System.Text.RegularExpressions;
namespace KNdata
{

	static class modMain
	{


		private const double RAD = 3.14159265358979 / 180.0;
			//TSのデータ格納Path
		static string oPath;
			//自分で管理するのデータファイル名
		static string[] oFile = new string[4];
			//TSのデータファイル名に付く、日時以外の文字列
		static string oFileA;
		static string tFile;
			//TSのデータから変位に
		static string heFile;

		static bool fUPDATE;

		static string LastFilename;

		static string LastDate;
		public struct zahyo
		{
			public int id;
			public double x;
			public double y;
			public double z;
		}

			//座標回転角度 DEG
		static double kakudo;

			//測点名称
		static string[] Pname = new string[15];
			//初期座標 元
		static zahyo[] INIT = new zahyo[15];
			//初期座標 回転後
		static zahyo[] dINIT = new zahyo[15];
			//変位の補正量 mm
		static zahyo[] offSET = new zahyo[15];


		static string[] sokutenName = new string[5];
		static int kanriLV;
		static double[] kanriV = new double[4];

		public static string LOGFILE;

		public static string ALERTfile;
		public static string[] KN_Path = new string[3];
		public static string[] KN_table = new string[4];
		public static string[] KN_Offset = new string[4];
		public static string[] KN_PathBK = new string[3];
		public static int[] sokutenSu = new int[4];

		public static string[] KN_SubName = new string[4];
		public static string[] SoushinPath = new string[4];

		public static string[] SoushinPathZ = new string[4];

		public static string[] GroupName = new string[4];

		public static string cuDir;

		public static void Main()
		{
			cuDir = KNdata.My.MyProject.Application.Info.DirectoryPath;
			//test
			cuDir = @"Y:\共有書庫\計測部員の書庫\白石書庫\業務\K_計測部\KN改";
			IniFile ini = new IniFile(cuDir + "\\TSdata.ini");

			int i = 0;
			int j = 0;

			KN_Path[1] = ini["system", "KN_Path1"];
			KN_Path[2] = ini["system", "KN_Path2"];
			if (Class1.strRight(KN_Path[1], 1) != "\\")
				KN_Path[1] = KN_Path[1] + "\\";
            if (Class1.strRight(KN_Path[2], 1) != "\\")
				KN_Path[2] = KN_Path[2] + "\\";

			KN_PathBK[1] = ini["system", "KN_MovePath1"];
			KN_PathBK[2] = ini["system", "KN_MovePath2"];
            if (Class1.strRight(KN_PathBK[1], 1) != "\\")
				KN_PathBK[1] = KN_PathBK[1] + "\\";
            if (Class1.strRight(KN_PathBK[2], 1) != "\\")
				KN_PathBK[2] = KN_PathBK[2] + "\\";

			KN_table[1] = ini["system", "KN_table1"];
			KN_table[2] = ini["system", "KN_table2"];
			KN_table[3] = ini["system", "KN_table3"];

			KN_Offset[1] = ini["system", "KN_Offset1"];
			KN_Offset[2] = ini["system", "KN_Offset2"];
			KN_Offset[3] = ini["system", "KN_Offset3"];

			SoushinPath[1] = ini["system", "SendPath1"];
			SoushinPath[2] = ini["system", "SendPath2"];
			SoushinPath[3] = ini["system", "SendPath3"];
            if (Class1.strRight(SoushinPath[1], 1) != "\\")
				SoushinPath[1] = SoushinPath[1] + "\\";
            if (Class1.strRight(SoushinPath[2], 1) != "\\")
				SoushinPath[2] = SoushinPath[2] + "\\";
            if (Class1.strRight(SoushinPath[3], 1) != "\\")
				SoushinPath[3] = SoushinPath[3] + "\\";

			SoushinPathZ[1] = ini["system", "SendPath1z"];
			SoushinPathZ[2] = ini["system", "SendPath2z"];
			SoushinPathZ[3] = ini["system", "SendPath3z"];
            if (Class1.strRight(SoushinPathZ[1], 1) != "\\")
				SoushinPathZ[1] = SoushinPathZ[1] + "\\";
            if (Class1.strRight(SoushinPathZ[2], 1) != "\\")
				SoushinPathZ[2] = SoushinPathZ[2] + "\\";
            if (Class1.strRight(SoushinPathZ[3], 1) != "\\")
				SoushinPathZ[3] = SoushinPathZ[3] + "\\";

            GroupName[1] = ini["Group", "Name1"];
			GroupName[2] = ini["Group", "Name2"];
			GroupName[3] = ini["Group", "Name3"];

			oPath = ini["system", "oPath"];
			oFile[1] = ini["system", "oFile1"];
			oFile[2] = ini["system", "oFile2"];
			oFile[3] = ini["system", "oFile3"];

			heFile = ini["system", "hFile"];
			oFileA = ini["system", "oFileA"];
			ALERTfile = ini["system", "ALERTfile"];

			sokutenSu[1] = 14;
			sokutenSu[2] = 13;
			sokutenSu[3] = 9;

			//1と2を入れ換えている
			KN_SubName[1] = "RAIL02";
			KN_SubName[2] = "RAIL01";

			//            Call ALERTfileCK("2017/09/19 20:00:00", j)

            double krv=0;
            string skrv = "";
            skrv = ini["kanri", "Vkanri1"];
            kanriV[1] = System.Convert.ToDouble(skrv);
            krv = System.Convert.ToDouble(ini["kanri", "Vkanri2"]);
            kanriV[2] = Convert.ToDouble(krv);
            krv = System.Convert.ToDouble(ini["kanri", "Vkanri3"]);
            kanriV[3] = Convert.ToDouble(krv);


			//"2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978

			LOGFILE = "TDdata.log";


			string st = "";
			int id = 0;

			id = 1;
			GetINIT(ref id);
			GetOffSet(ref id);
			LastDate = sLastDate(ref id);
			//自分が管理するファイルの最終日時
			//    Debug.Print DTMtoFname(LastDate)
			//    LastFilename = KN_Path(id) & "R" & DTMtoFname(LastDate) & KN_SubName(id) & "_Total.txt"
			LastFilename = "R" + DTMtoFname(ref LastDate) + KN_SubName[id] + "_Total.txt";

			// WriteLog id & ":" & LastFilename
			//    kanrihantei 1, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"


			int co = 0;
			string[] tsFile = {
				
			};
			co = CheckDataFile(KN_Path[id], ref tsFile);

			// WriteLog id & ":" & co

			if (0 < co) {
				AppendData(id, ref tsFile);
				if (Exists2(cuDir + "\\fSoushin.exe") == true) {
                    System.Diagnostics.Process.Start(cuDir + "\\fSoushin.exe");

				}
			}

			id = 2;
			GetINIT(ref id);
			GetOffSet(ref id);
			LastDate = sLastDate(ref id);
			//自分が管理するファイルの最終日時
			//    Debug.Print DTMtoFname(LastDate)
			LastFilename = "R" + DTMtoFname(ref LastDate) + KN_SubName[id] + "_TOTAL.TXT";
			//kanrihantei id, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"

			co = 0;
			tsFile = null;
			co = CheckDataFile(KN_Path[id], ref tsFile);

			if (0 < co) {
				AppendData(id, ref tsFile);
				if (Exists2(cuDir + "\\fSoushin.exe") == true) {
                    System.Diagnostics.Process.Start(cuDir + "\\fSoushin.exe");
                }

                id = 3;
                GetINIT(ref id);
                GetOffSet(ref id);
                LastDate = sLastDate(ref id);
                //自分が管理するファイルの最終日時
                //    Debug.Print DTMtoFname(LastDate)
                //ID=2 と ID=3 は同じファイルをみる
                LastFilename = "R" + DTMtoFname(ref LastDate) + KN_SubName[2] + "_TOTAL.TXT";
                //kanrihantei id, "2017/08/26 09:00:00, -35, -2.32, 41, -0.24, 0.37, 31, -0.4, 0.25, 0, -0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 120,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2,-0.15, 3.2, 0.2"

                co = 0;
                tsFile = null;
                co = CheckDataFile(KN_Path[2], ref tsFile);

                if (0 < co)
                {
                    AppendData(id, ref tsFile);
                    if (Exists2(cuDir + "\\fSoushin.exe") == true)
                    {
                        System.Diagnostics.Process.Start(cuDir + "\\fSoushin.exe");
                    }
                }            
			}
		}


		private static void GetINIT(ref int id)
		{
			string[] sa = null;
			string[] sb = null;
			string bf = "";
			int i = 0;
			int j = 0;
            string[] del = {"\r\n"};

            try {
				using (StreamReader sr = new StreamReader(KN_table[id], Encoding.GetEncoding("Shift_JIS"))) {
					bf = sr.ReadToEnd();
					sr.Close();
				}

				//Console.Write(bf)

                sa = bf.Split(del, StringSplitOptions.None);
				for (i = 0; i <= sa.Length; i++) {
					if (!string.IsNullOrEmpty(sa[i])) {
						switch (sa[i].Substring(0, 1)) {
							case ";":
							case ":":
							case "'":
								break;
							default:
                                sb = sa[i].Split(',');
								j = Convert.ToInt32(sb[0]);
								Pname[j] = sb[1];
								INIT[j].x = Convert.ToDouble(sb[2]);
								INIT[j].y = Convert.ToDouble(sb[3]);
								INIT[j].z = Convert.ToDouble(sb[4]);
								break;
						}
					}
				}

			} catch (Exception exception) {
				//Console.WriteLine(exception.Message)
			}

		}

		private static void GetOffSet(ref int id)
		{
			string[] sa = null;
			string[] sb = null;
			string bf = null;
			int i = 0;
			int j = 0;
            string[] del = { "\r\n" };

			try {
				using (StreamReader sr = new StreamReader(KN_Offset[id], Encoding.GetEncoding("Shift_JIS"))) {
					bf = sr.ReadToEnd();
					sr.Close();
				}

                sa = bf.Split(del, StringSplitOptions.None);
				for (i = 0; i <= sa.Length; i++) {
					if (!string.IsNullOrEmpty(sa[i])) {
						switch (sa[i].Substring(0, 1)) {
							case ";":
							case ":":
							case "'":
								break;
							default:
								sb = sa[i].Split(',');
								j = Convert.ToInt32(sb[0]);
								//                Pname(j) = sb(1)
								offSET[j].x = Convert.ToDouble(sb[2]);
								offSET[j].y = Convert.ToDouble(sb[3]);
								offSET[j].z = Convert.ToDouble(sb[4]);
								break;
						}
					}
				}

			} catch (Exception exception) {
				//Console.WriteLine(exception.Message)
			}
		}

		public static void zahyohenkan(ref zahyo[] dt)
		{
			double a11 = 0;
			double a12 = 0;
			double a21 = 0;
			double a22 = 0;

			a11 = System.Math.Cos(kakudo * RAD);
			a12 = System.Math.Sin(kakudo * RAD);
			a21 = -System.Math.Sin(kakudo * RAD);
			a22 = System.Math.Cos(kakudo * RAD);

			int i = 0;
			double x = 0;
			double y = 0;
			double xx = 0;
			double yy = 0;

			for (i = 1; i <= dt.Length; i++) {
				x = dt[i].x;
				y = dt[i].y;
				xx = a11 * x + a12 * y;
				yy = a21 * x + a22 * y;

				dt[i].x = xx;
				dt[i].y = yy;
			}


		}


		public static int CheckDataFile(string fdir, ref string[] tFile)
		{
			int functionReturnValue = 0;

			int lIndex = 0;

			int i = 0;
			int j = 0;

			//string ret = "";
            string fFilename = "";
			string[] tFilename = {};
			int aIndex = 0;
			aIndex = -1;

			lIndex = 0;

			foreach (string filepath in Directory.GetFiles(fdir, "*.txt", SearchOption.TopDirectoryOnly)) {
                fFilename = modFiles.FindFileName(filepath);
                if (fFilename.Substring(0,1) == "R" && fFilename.Substring(fFilename.Length-3, 3) == "TXT") {
					lIndex = lIndex + 1;
					aIndex = aIndex + 1;
					Array.Resize(ref tFilename, aIndex + 1);
					tFilename[aIndex] = modFiles.FindFileName(filepath);
				}
				//            Console.WriteLine(filepath)
			}

			functionReturnValue = lIndex;
			if (lIndex == 0) {
				functionReturnValue = 0;
				return functionReturnValue;
			}

			//所得したファイル名をソート
			if (-1 < aIndex) {
				modFiles.s_ShellSort(ref tFilename, (aIndex));
			}

            string az=null;
            string bz=null;

			for (i = 0; i <= aIndex; i++) {
                az = LastFilename.ToUpper();
                bz = tFilename[i].ToUpper();
                if (az.CompareTo( bz ) < 0) {
					j = j + 1;
					Array.Resize(ref tFile, j + 1);
					tFile[j] = tFilename[i];
				}
			}
			functionReturnValue = j;
			return functionReturnValue;

		}


		public static string FnametoDTM(string st)
		{
			string functionReturnValue = null;
			//TSファイル名から日時を生成
			//st : TSファイル名 20170826_09

			 // ERROR: Not supported in C#: OnErrorStatement

			st = st.Replace(".txt", "");
			st = st.Replace(".TXT", "");
            st = st.ToUpper();
            st = st.Replace("_TOTAL", "");
			string sst = null;
			string yy = null;
			string mm = null;
			string dd = null;
			string hh = null;
			string nn = null;
			string ss = null;

			yy = Class1.strMid(st, 2, 4);
            mm = Class1.strMid(st, 6, 2);
            dd = Class1.strMid(st, 8, 2);
            hh = Class1.strMid(st, 10, 2);
            nn = Class1.strMid(st, 12, 2);
            ss = Class1.strMid(st, 14, 2);

			DateTime dt = new DateTime(Convert.ToInt32(yy), Convert.ToInt32(mm), Convert.ToInt32(dd), Convert.ToInt32(hh), 0, 0, DateTimeKind.Local);
			sst = dt.ToString("yyyy/MM/dd HH:mm:ss");

			functionReturnValue = sst;
			 // ERROR: Not supported in C#: OnErrorStatement

			return functionReturnValue;

            /*
            FnametoDTM9999:
			functionReturnValue = "";
			 // ERROR: Not supported in C#: OnErrorStatement

			return functionReturnValue;
             */
		}

		public static string DTMtoFname(ref string st)
		{
			string functionReturnValue = null;
			//日時フォーマットからファイル名を生成
			//st : 日時フォーマット
			DateTime dt = default(DateTime);
			string sst = null;
			if (DateTime.TryParse(st, out dt) == true) {
				sst = dt.ToString("yyyyMMddHHmmss");
				functionReturnValue = sst;
			} else {
				functionReturnValue = "";
			}
			return functionReturnValue;
		}

		public static string DTMtoDname(ref string st)
		{
			string functionReturnValue = null;
			//日時フォーマットからディレクトリ名を生成
			//st : 日時フォーマット
			DateTime dt = default(DateTime);
			string sst = null;
			if (DateTime.TryParse(st, out dt) == true) {
				sst = dt.ToString("yyyy-MM");
				functionReturnValue = sst;
			} else {
				functionReturnValue = "";
			}
			return functionReturnValue;
		}

		public static string sLastDate(ref int id)
		{
			string functionReturnValue = null;
			//保存データファイルの最終日時を取得する
			// ID : データ番号
			// ed : 最終日時

			 // ERROR: Not supported in C#: OnErrorStatement


			string nm = oFile[id];
			string ed = "";

			System.IO.FileInfo fi = new System.IO.FileInfo(nm);
			//ファイルのサイズを取得
			long l = fi.Length;
			long sl = 0;
			long sp = 0;
			sl = l;
			do {
				sl = sl / 2;
				if (sl < 1024) {
					sp = l - sl;
					FileStream fs = null;
					StreamReader sr = null;
					char[] buf = new char[3];

					fs = new FileStream(nm, FileMode.Open, FileAccess.Read);
					sr = new StreamReader(fs);
					fs.Seek(sp, SeekOrigin.Begin);
					//sr.ReadBlock(buf, 0, buf.Length)

					while (-1 < sr.Peek()) {
						ed = sr.ReadLine();
					}

					sr.Close();
					fs.Close();
					ed = Class1.strMid(ed, 1, 19);
					functionReturnValue = ed;
					break; // TODO: might not be correct. Was : Exit Do
				}
			} while (true);
			 // ERROR: Not supported in C#: OnErrorStatement

			return functionReturnValue;

            /*
            LastDate9999:

			functionReturnValue = "";
			 // ERROR: Not supported in C#: OnErrorStatement

			return functionReturnValue;
             */
		}


		public static void AppendData(int id, ref string[] fNam)
		{
			string n1 = null;
			string bf = null;
			string wbf = "";

			int ii = 0;
			int i = 0;
			int j = 0;
			string[] sa = null;
			string[] sb = null;

			string MDY = null;
			zahyo[] dt = new zahyo[15];
			zahyo[] heniDT = new zahyo[15];
			zahyo[] heni = new zahyo[15];
			int cc = 0;
			int no = 0;
			bool fx = false;
			int tID = 0;

			for (i = 0; i <= 14; i++) {
				dt[i].x = 999999;
				dt[i].y = 999999;
				dt[i].z = 999999;
				heniDT[i].x = 999999;
				heniDT[i].y = 999999;
				heniDT[i].z = 999999;
				heni[i].x = 999999;
				heni[i].y = 999999;
				heni[i].z = 999999;
			}

			if (id == 3) {
				tID = 2;
			} else {
				tID = id;
			}

			System.IO.StreamWriter swMS = new System.IO.StreamWriter(oFile[id], true, System.Text.Encoding.GetEncoding("shift_jis"));

            string[] del = {"\r\n"};

			for (ii = 1; ii <= fNam.Length; ii++) {
				//    WriteLog KN_Path(tID) & UCase(fNam(ii))

				n1 = fNam[ii].ToUpper();
				if (("R" + DTMtoFname(ref LastDate).ToUpper() + KN_SubName[tID].ToUpper() + "_TOTAL.TXT").CompareTo(n1) == -1) {
					if (System.IO.File.Exists(KN_Path[tID] + n1) == false) {
						return;
					}

					fx = true;

					System.IO.StreamReader swTS = new System.IO.StreamReader(KN_Path[tID] + n1, System.Text.Encoding.GetEncoding("shift_jis"));
					//ファイル全体を読み込み
					bf = swTS.ReadToEnd();
					//オープンしていたファイルを閉じる
					swTS.Close();

					MDY = FnametoDTM(n1);

					sa = bf.Split(del, StringSplitOptions.None );
					for (i = 0; i <= sa.Length; i++) {
						if (!string.IsNullOrEmpty(sa[i])) {
							sb = sa[i].Split(',');
							for (j = 1; j <= sokutenSu[id]; j++) {
								if ((sb[1].ToUpper() + sb[2].ToUpper()) == Pname[j]) {
									if (sb[3] == "0" && sb[4] == "0") {
										no = j;
										//sb(2)
										dt[no].x = Convert.ToDouble(sb[14]);
										dt[no].y = Convert.ToDouble(sb[16]);
										dt[no].z = Convert.ToDouble(sb[18]);
										heniDT[no].x = Convert.ToDouble(sb[6]);
										heniDT[no].y = Convert.ToDouble(sb[8]);
										heniDT[no].z = Convert.ToDouble(sb[10]);
										//                                Debug.Print Pname(j), j
									}
									break; // TODO: might not be correct. Was : Exit For
								}
							}
						}
					}
					//    Debug.Print sa(0)
					wbf = MDY;
					for (i = 1; i <= sokutenSu[id]; i++) {
						wbf = wbf + "," + dt[i].x + "," + dt[i].y + "," + dt[i].z;
					}
					swMS.WriteLine(wbf);

					DateTime dte = default(DateTime);

					DateTime.TryParse(MDY, out dte);

					System.IO.StreamWriter swSS = new System.IO.StreamWriter(SoushinPathZ[id] + dte.ToString("yyyy-MM-dd_HH-mm-ss") + ".csv", false, System.Text.Encoding.GetEncoding("shift_jis"));

					wbf = MDY;
					for (i = 1; i <= sokutenSu[id]; i++) {
						wbf = wbf + "," + dt[i].x.ToString("0.0000") + "," + dt[i].y.ToString("0.0000") + "," + dt[i].z.ToString("0.0000");
					}
					swSS.WriteLine((wbf));
					swSS.Close();
					//変位量 (mm)
					for (i = 1; i <= sokutenSu[id]; i++) {
						if (heniDT[i].x == 999999) {
							heni[i].x = 999999;
						} else {
							heni[i].x = (heniDT[i].x - INIT[i].x) - offSET[i].x;
						}
						if (heniDT[i].y == 999999) {
							heni[i].y = 999999;
						} else {
							heni[i].y = (heniDT[i].y - INIT[i].y) - offSET[i].y;
						}
						if (heniDT[i].z == 999999) {
							heni[i].z = 999999;
						} else {
							heni[i].z = (heniDT[i].z - INIT[i].z) - offSET[i].z;
						}
					}

					System.IO.StreamWriter swSS2 = new System.IO.StreamWriter(SoushinPath[id] + dte.ToString("yyyy-MM-dd_HH-mm-ss") + ".csv", false, System.Text.Encoding.GetEncoding("shift_jis"));

                    string fmt = "0.0000";
                    wbf = MDY;
					for (i = 1; i <= sokutenSu[id]; i++) {
						wbf = wbf + "," + FormatD(ref heni[i].x, ref fmt) + "," + FormatD(ref heni[i].y, ref fmt) + "," + FormatD(ref heni[i].z, ref fmt);
					}
					swSS2.WriteLine((wbf));
					swSS2.Close();

					System.IO.StreamWriter swSS3 = new System.IO.StreamWriter(cuDir + "\\Newest" + id + ".csv", false, System.Text.Encoding.GetEncoding("shift_jis"));
					wbf = MDY;
					for (i = 1; i <= sokutenSu[id]; i++) {
						wbf = wbf + "," + FormatD(ref heni[i].x, ref fmt) + "," + FormatD(ref heni[i].y, ref fmt) + "," + FormatD(ref heni[i].z, ref fmt);
					}
					swSS3.WriteLine((wbf));
					swSS3.Close();

				}
				if (id != 2) {
					DoFileMove(KN_Path[tID] + n1, KN_PathBK[tID] + n1);
				}
			}
			//閉じる
			swMS.Close();

			if (fx == true) {
				kanrihantei(ref id, ref wbf);
			}

		}

		public static string FormatD(ref double dt, ref string fmt)
		{
			string functionReturnValue = null;
			if (System.Math.Abs(dt) == 999999) {
				functionReturnValue = "999999";
			} else {
				functionReturnValue = dt.ToString(fmt);
			}
			return functionReturnValue;
		}

		//2017/08/26 09:00:00,-0.359807621135744,-2.32050807564832E-02,0.199999999999978,-0.2464101615125,0.373205080755668,0.099999999999989,-0.433012701890334,0.249999999999417,0,-0.15980762113621,0.323205080757116,0.199999999999978
		public static void kanrihantei(ref int id, ref string bf)
		{
			string[] sa = null;
			int i = 0;

			double[] xd = new double[15];
			double[] yd = new double[15];
			double[] zd = new double[15];

			kanriLV = -1;
			sa = bf.Split(',');
			for (i = 1; i <= (sa.Length); i++) {
				switch ((i % 3)) {
					case 1:
						xd[(i / 3) + 1] = Convert.ToDouble(sa[i]);
						break;
					case 2:
						yd[(i / 3) + 1] = Convert.ToDouble(sa[i]);
						break;
					case 0:
						zd[(i / 3) + 0] = Convert.ToDouble(sa[i]);
						break;
				}
			}

			//管理レベルを調べる
			for (i = 1; i <= (sa.Length) / 3; i++) {
				if (zd[i] != 999999) {
					if (!(-kanriV[3] < zd[i] & zd[i] < kanriV[3])) {
						kanriLV = 3;
					} else if (!(-kanriV[2] < zd[i] & zd[i] < kanriV[2])) {
						if (kanriLV < 3)
							kanriLV = 2;
					} else if (!(-kanriV[1] < zd[i] & zd[i] < kanriV[1])) {
						if (kanriLV < 2)
							kanriLV = 1;
					} else {
						if (kanriLV < 1)
							kanriLV = 0;
					}
				}
			}

			int ret = 0;
			string sda = null;
			sda = sa[0];
			if (0 < kanriLV) {
				if (Exists2(cuDir + "\\ALARMsw.exe")) {
					WriteLog("アラームSW　ON " + kanriLV);
					//            Call ALERTfileCK(sda, ret)
					if (ret < 0) {
                        System.Diagnostics.Process.Start(cuDir + "\\ALARMsw.exe " + kanriLV);
					} else if (ret < kanriLV) {
                        System.Diagnostics.Process.Start(cuDir + "\\ALARMsw.exe " + kanriLV);
					} else {
                        System.Diagnostics.Process.Start(cuDir + "\\ALARMsw.exe " + ret);
					}
				}
			}

			if (kanriLV <= 0) {
				return;
			}

			string alst = null;

			alst = "以下の計測データが管理値を超過しました。" + "\r\n";
			alst = alst + "\r\n" + "計測日時：" + sa[0];

			//管理レベルを超えた測点を調べる
			for (i = 1; i <= (sa.Length) / 3; i++) {
				if (zd[i] != 999999) {
					if (!(-kanriV[3] < zd[i] & zd[i] < kanriV[3])) {
						//                alst = alst & vbCrLf & "管理レベルⅢ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
                        alst = alst + "\r\n" + "管理レベルⅢ超過 : " + NameCHG(ref Pname[i]) + " 沈下量 = " + zd[i] + " mm";
					} else if (!(-kanriV[2] < zd[i] & zd[i] < kanriV[2])) {
						//                alst = alst & vbCrLf & "管理レベルⅡ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
                        alst = alst + "\r\n" + "管理レベルⅡ超過 : " + NameCHG(ref Pname[i]) + " 沈下量 = " + zd[i] + " mm";
					} else if (!(-kanriV[1] < zd[i] & zd[i] < kanriV[1])) {
						//                alst = alst & vbCrLf & "管理レベルⅠ超過 : " & GroupName(id) & i & " Z方向 = " & zd(i)
                        alst = alst + "\r\n" + "管理レベルⅠ超過 : " + NameCHG(ref Pname[i]) + " 沈下量 = " + zd[i] + " mm";
					}
				}
			}

			alst = alst + "\r\n" + "========================================";

            File.AppendAllText(cuDir + "\\send0000.txt", alst);

			if (Exists2(cuDir + "\\kmSoushin.exe") == true) {
                System.Diagnostics.Process.Start(cuDir + "\\kmSoushin.exe");
			}

		}

        public static void WriteLog(string st)
		{
			//st 説明文
			 // ERROR: Not supported in C#: OnErrorStatement

            File.AppendAllText(cuDir + "\\" + LOGFILE, DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " : " + Environment.NewLine);
            File.AppendAllText(cuDir + "\\" + LOGFILE, st  + Environment.NewLine);

			 // ERROR: Not supported in C#: OnErrorStatement

		}

		public static string SeName(ref int id, ref string da)
		{
			string functionReturnValue = null;
			int p1 = 0;
			int p2 = 0;
			p1 = da.IndexOf('/');
			p2 = da.IndexOf('/',p1 + 1);
			if (p1 == 0 | p2 == 0) {
				functionReturnValue = "";
			}

			string sYY = null;
			string sMM = null;

			sYY = Class1.strMid(da, 1, p1 - 1);
            sMM = Class1.strMid(da, p1 + 1, p2 - p1 - 1);
			functionReturnValue = id + "__" + sYY + sMM;
			return functionReturnValue;
		}

		////// ファイル及びフォルダの有無チェック(有無のみ判定) /////
		// True=存在する、False=存在しない
		//-----------------------------------------------------------
		public static bool Exists2(string strPathName)
		{
			bool functionReturnValue = false;
			//strPathName : フルパス名
			//------------------------
			 // ERROR: Not supported in C#: OnErrorStatement
            if (System.IO.File.Exists(strPathName))
            {
                functionReturnValue = true;
                return functionReturnValue;
            }

            if (System.IO.Directory.Exists(strPathName))
            {
                functionReturnValue = true;
                return functionReturnValue;
            }
            
			//Debug.Print strPathName & "が見つかりません。"
			 // ERROR: Not supported in C#: OnErrorStatement

			return functionReturnValue;
		}

		public static void test()
		{
		}

		public static void DoFileMove(string sp, string dp)
		{
			//sp:元
			//dp:先
			Scripting.FileSystemObject Fso = new Scripting.FileSystemObject();

			string ssp = null;
			string sdp = null;
			ssp = Fso.GetAbsolutePathName(sp);
			sdp = Fso.GetAbsolutePathName(dp);

			string dDirectory = null;
			string fNam = null;
			string pa = null;
			dDirectory = Fso.GetParentFolderName(sdp);
			//; // "C:\\data" が返る
			fNam = Fso.GetFileName(sdp);
			pa = FnametoDTM(fNam);

			MakeDirectory(dDirectory + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(pa, "\\yyyy\\MM\\dd"));
			sdp = dDirectory + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(pa, "\\yyyy\\MM\\dd") + "\\" + fNam;
			Fso.CopyFile(ssp, sdp, true);
			Fso.DeleteFile(ssp, true);

		}

		public static void MakeDirectory(string sPath)
		{
			//深い階層のディレクトリまで作成
			System.IO.Directory.CreateDirectory(sPath);
		}

		//以下 2017年11月22日 追加
		private static string NameCHG(ref string st)
		{
			string st0 = null;
			string st1 = null;
			string st2 = null;
			st1 = st.Substring(0, 1);
			st0 = st1 + "-";
			st2 = st.Substring(1,st.Length - 1);
			st1=FindNumberRegExp(ref st2);
			return st0 + st1;
		}

		//// 引数1：対象文字列
		//// 引数2：検索結果
        private static string FindNumberRegExp(ref string s)
		{
            string result = "";
            if (s.IndexOf("0") == 0) {
                result = s;
                return result;
			}
			//Dim result As Boolean = Regex.IsMatch("{検査対象文字列}", "{正規表現パターン}")
			Match Reslt = Regex.Match(s, "[0-9]");
            return result;
		}

    }

}  
