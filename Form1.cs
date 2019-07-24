using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;

namespace MeasureStatusMonitor
{
    public partial class Form1 : Form
    {
//PCがスリープ状態に入らないようする start
        #region Win32 API
        [FlagsAttribute]
        public enum ExecutionState : uint
        {
            // 関数が失敗した時の戻り値
            Null = 0,
            // スタンバイを抑止(Vista以降は効かない？)
            SystemRequired = 1,
            // 画面OFFを抑止
            DisplayRequired = 2,
            // 効果を永続させる。ほかオプションと併用する。
            Continuous = 0x80000000,
        }

        [DllImport("user32.dll")]
        extern static uint SendInput(
            uint nInputs,   // INPUT 構造体の数(イベント数)
            INPUT[] pInputs,   // INPUT 構造体
            int cbSize     // INPUT 構造体のサイズ
            );

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct INPUT
        {
            public int type;  // 0 = INPUT_MOUSE(デフォルト), 1 = INPUT_KEYBOARD
            public MOUSEINPUT mi;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;  // amount of wheel movement
            public int dwFlags;
            public int time;  // time stamp for the event
            public IntPtr dwExtraInfo;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        // dwFlags
        const int MOUSEEVENTF_MOVED = 0x0001;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;  // 左ボタン Down
        const int MOUSEEVENTF_LEFTUP = 0x0004;  // 左ボタン Up
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;  // 右ボタン Down
        const int MOUSEEVENTF_RIGHTUP = 0x0010;  // 右ボタン Up
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;  // 中ボタン Down
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;  // 中ボタン Up
        const int MOUSEEVENTF_WHEEL = 0x0080;
        const int MOUSEEVENTF_XDOWN = 0x0100;
        const int MOUSEEVENTF_XUP = 0x0200;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        const int screen_length = 0x10000;  // for MOUSEEVENTF_ABSOLUTE
        [DllImport("kernel32.dll")]
        static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);
        #endregion
//PCがスリープ状態に入らないようする end



        static public SYSSET SETDATA = new SYSSET();
		public string currentCsvFile = "";
		public string InfoStr = "";

		public string operatorName = "";
		public string seikeikiNo = "";
		public string goukiNo = "";
		public string hinshu = "";
		public string sleeveNo = "";
		public string cavNo = "";
		public string priority = "";
		public string dateInfo = "";
		public string timeInfo = "";
		public string tairyuTime = "";
		public string status = "";
		public string endDate = "";
		public string endTime = "";
		public bool isUpdate = false;
		public DateTime targetDate;
        List<Color> colorList = new List<Color>();

        public Form1()
        {
            InitializeComponent();
            SETDATA.load(ref Form1.SETDATA);

            this.Width = SETDATA.windowWidth;
			this.Height = SETDATA.windowHeight;

			listView1.Width = SETDATA.listviewWidth;
			listView1.Height = SETDATA.listviewHeight;

			listView1.FullRowSelect = true;
			listView1.GridLines = true;
			listView1.Sorting = SortOrder.Ascending;
			listView1.View = View.Details;
			listView1.HideSelection = false;

			// 列（コラム）ヘッダの作成
			ColumnHeader columnOperator;
			ColumnHeader columnSeikeiki;
			ColumnHeader columnGouki;
			ColumnHeader columnHinshu;
			ColumnHeader columnSleeveNo;
			ColumnHeader columnCavNo;
			ColumnHeader columnDate;
			ColumnHeader columnTime;
			ColumnHeader columnTairyu;
			ColumnHeader columnStatus;
			ColumnHeader columnEndDate;
			ColumnHeader columnEndTime;

			columnOperator = new ColumnHeader();
			columnSeikeiki = new ColumnHeader();
			columnGouki = new ColumnHeader();
			columnHinshu = new ColumnHeader();
			columnSleeveNo = new ColumnHeader();
			columnCavNo = new ColumnHeader();
			columnDate = new ColumnHeader();
			columnTime = new ColumnHeader();
			columnTairyu = new ColumnHeader();
			columnStatus = new ColumnHeader();
			columnEndDate = new ColumnHeader();
			columnEndTime = new ColumnHeader();

            listView1.Sorting = SortOrder.None;
            listView1.ForeColor = Color.Black;//初期の色
            listView1.BackColor = Color.White;//背景色

			columnDate.Text = "登録日";
			columnTime.Text = "登録時間";
			columnOperator.Text = "登録者";
			columnSeikeiki.Text = "成型機";
			columnGouki.Text = "号機";
			columnHinshu.Text = "製品品種";
			columnSleeveNo.Text = "スリーブNo";
			columnCavNo.Text = "CavNo";
			columnTairyu.Text = "滞留時間";
			columnStatus.Text = "状態";
			columnEndDate.Text = "終了日";
			columnEndTime.Text = "終了時間";

			ColumnHeader[] colHeaderRegValue = {columnDate, columnTime, columnOperator, columnSeikeiki, columnGouki, columnHinshu, columnSleeveNo, columnCavNo, columnTairyu, columnStatus, columnEndDate, columnEndTime};
			listView1.Columns.AddRange(colHeaderRegValue);

			//ヘッダの幅を自動調節
            listView1.Font = new System.Drawing.Font("Times New Roman", 16, System.Drawing.FontStyle.Regular);
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);


	        // 出力用のファイルを開く
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
            currentCsvFile = stCurrentDir + "\\PgmMeasureInfo" + ".csv";
			ReadCurrentCSV();

			timer1.Interval = SETDATA.timerOut * 1000;
			timer1.Enabled = true;
			label1.Text = string.Format("{0}時間未満", SETDATA.hourLimit);
			label2.Text = "測定終了";
			label3.Text = string.Format("{0}時間以上　経過", SETDATA.hourLimit);
			label4.Text = string.Format("{0}日　以上　経過", SETDATA.dayLimit);
            label10.Text = "最優先";
            label12.Text = "　優先";
			
			int openCount = 0;
			int greenCount = 0;
			int orangeCount = 0;
			int magentaCount = 0;
			int redCount = 0;
			int yellowCount = 0;
			for(int i = 0; i < listView1.Items.Count; i++)
			{
				if(listView1.Items[i].SubItems[9].Text == "測定待ち")
				{
					openCount++;
				}

				if(listView1.Items[i].BackColor == Color.Lime)
				{
					greenCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Orange)
				{
					orangeCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Magenta)
				{
					magentaCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Red)
				{
					redCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Yellow)
				{
					yellowCount++;
				}
			}
			label5.Text = string.Format("{0}", openCount);
			
			label6.Text = string.Format("{0}", greenCount);
			label7.Text = string.Format("{0}", orangeCount);
			label8.Text = string.Format("{0}", magentaCount);
			label11.Text = string.Format("{0}", redCount);
			label13.Text = string.Format("{0}", yellowCount);

			if(redCount == 0)
			{
				timer5.Enabled = false;
				label10.BackColor = Color.Red;
				label11.BackColor = Color.Red;
				label9.BackColor = Color.White;
			}
			else
			{
				timer5.Enabled = true;
				label9.BackColor = Color.Red;
			}

			if(yellowCount == 0)
			{
				timer6.Enabled = false;
				label12.BackColor = Color.Yellow;
				label13.BackColor = Color.Yellow;
				label9.ForeColor = Color.Black;
			}
			else
			{
				timer6.Enabled = true;
				label9.ForeColor = Color.Yellow;
			}

			//ログに出力する
			StatusFileOut(string.Format("{0}", openCount));

			timer2.Enabled = true;//画面暗転阻止
			timer3.Enabled = true;//現在時刻


			targetDate = new DateTime(2019, 4, 1, 0, 0, 0);

			System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
			System.Version ver = asm.GetName().Version;
            this.Text += "  Ver:" + ver;
        }

		public void ReadCurrentCSV()
		{
			//ListViewを一度クリアする
			listView1.Items.Clear();
			colorList.Clear();
			
			string[] item1 = {dateInfo, timeInfo, operatorName, seikeikiNo, goukiNo, hinshu, sleeveNo, cavNo, tairyuTime, status, endDate, endTime, priority};

            try
            {
	            //CSV読込。高速化を狙い、最大行数を取得後、for分でループする
	            var readToEnd = File.ReadAllLines(@currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
	            int lines = readToEnd.Length;

	            for (int i = 0; i < lines; i++)
	            {
					if(i == 0)//ヘッダ部はスキップ
					{
						continue;
					}
	                //１行のstringをstream化してTextFieldParserで処理する
	                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
	                {
	                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
	                    {
	                        parser.TextFieldType = FieldType.Delimited;
	                        parser.Delimiters = new[] { "," };
	                        parser.HasFieldsEnclosedInQuotes = true;
	                        parser.TrimWhiteSpace = false;
	                        string[] fields = parser.ReadFields();

	                        for (int j = 0; j < fields.Length; j++)
	                        {
								item1[j] = fields[j];
	                        }
	                    }
	                }

					listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加

					if(item1[12] == "最優先")
					{
						listView1.Items[0].BackColor = Color.Red;//背景色
						colorList.Insert(0, Color.Red);
					}
					else if(item1[12] == "優先")
					{
						listView1.Items[0].BackColor = Color.Yellow;//背景色
						colorList.Insert(0, Color.Yellow);
					}
					else//通常
					{
						if(item1[9] == "測定終了")
						{
							listView1.Items[0].BackColor = Color.Gray;//背景色
							colorList.Insert(0, Color.Gray);
						}
						else
						{
							string strTime = item1[0] + " " + item1[1];
							DateTime dTime = DateTime.Parse(strTime);
							DateTime dt = DateTime.Now;
							TimeSpan ts = dt - dTime;

	                        if(ts.Days < SETDATA.dayLimit)
	                        {
								if(0 < ts.Days)
								{
									listView1.Items[0].BackColor = Color.Orange;//背景色
									colorList.Insert(0, Color.Orange);
								}
								else
								{
									if(ts.Hours < SETDATA.hourLimit)
									{
										listView1.Items[0].BackColor = Color.Lime;//背景色
										colorList.Insert(0, Color.Lime);
									}
		                            else
									{
										listView1.Items[0].BackColor = Color.Orange;//背景色
										colorList.Insert(0, Color.Orange);
									}
								}
							}
							else
							{
								listView1.Items[0].BackColor = Color.Magenta;//背景色
								colorList.Insert(0, Color.Magenta);
							}
						}
					}

	            }
			}
			catch (System.IO.IOException ex)
		    {
		        // ファイルを開くのに失敗したときエラーメッセージを表示
				string errorStr = "状変時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			listView1.Font = new System.Drawing.Font("Times New Roman", 18, System.Drawing.FontStyle.Regular);
            //ヘッダの幅を自動調節
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
		}

        public void LogFileOut(string logMessage)
		{
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string path = stCurrentDir + "\\PgmMeasureInfo.log";
			
			using(var sw = new System.IO.StreamWriter(path, true, System.Text.Encoding.Default))
			{
				sw.WriteLine($"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()}");
				sw.WriteLine($"  {logMessage}");
				sw.WriteLine ("--------------------------------------------------------------");
			}
		}

        public void StatusFileOut(string logMessage)
		{
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string path = stCurrentDir + "\\PgmStatusInfo.log";
			
			using(var sw = new System.IO.StreamWriter(path, true, System.Text.Encoding.Default))
			{
				sw.WriteLine($"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()} {logMessage}");
			}
		}

		public bool WriteDataToCsv(string logStr)
		{
            try
            {
		        // appendをtrueにすると，既存のファイルに追記
		        //         falseにすると，ファイルを新規作成する
		        var append = false;
		        // 出力用のファイルを開く
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                currentCsvFile = stCurrentDir + "\\PgmMeasureInfo" + ".csv";

                string buf = "";
                if(System.IO.File.Exists(currentCsvFile))//既にファイルが存在する
				{
					append = true;
				}
				
		        using(var sw = new System.IO.StreamWriter(currentCsvFile, append, System.Text.Encoding.Default))
		        {
					if(!append)
					{
						buf = string.Format("登録日");
						buf += string.Format(",登録時間");
						buf += string.Format("登録者");
                        buf += string.Format(",成型機");
                        buf += string.Format(",号機");
                        buf += string.Format(",製品品種");
						buf += string.Format(",スリーブNo");
						buf += string.Format(",CavNo");
						buf += string.Format(",滞留時間");
						buf += string.Format(",状態");
						buf += string.Format(",終了日");
						buf += string.Format(",終了時間");
						buf += string.Format(",優先度");

	                    sw.WriteLine(buf);

						buf = "";
						buf += logStr;
	                    sw.WriteLine(buf);
					}
					else
					{
						buf = "";
						buf += logStr;
	                    sw.WriteLine(buf);
					}
		        }
		    }
			catch (System.IO.IOException ex)
		    {
		        // ファイルを開くのに失敗したときエラーメッセージを表示
				string errorStr = "状変時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
				return false;
		    }
		    return true;

		}

		public class SYSSET:System.ICloneable
		{
			public int windowWidth;
			public int windowHeight;
			public int listviewHeight;
			public int listviewWidth;

			public int timerOut;
			public int dayLimit;
			public int hourLimit;

			public string[] machineKind = {"", "", "", ""};

			public string[] seihinName=		{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
											 "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};

			public string[] registerOperator = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
												"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
												"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
												"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
												"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""};
			
	        public bool load(ref SYSSET ss)
			{
	            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
	            string path = stCurrentDir + "\\MeasureStatusSetting.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamReader fs = new System.IO.StreamReader(path, System.Text.Encoding.Default);
					SYSSET obj;
					obj = (SYSSET)sz.Deserialize(fs);
					fs.Close();
					obj = (SYSSET)obj.Clone();
					ss = obj;
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return(ret);
			}

			public Object Clone()
			{
				SYSSET cln = (SYSSET)this.MemberwiseClone();
				return (cln);
			}

			public bool save(SYSSET ss)
			{
	            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
	            string path = stCurrentDir + "\\MeasureStatusSetting.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamWriter fs = new System.IO.StreamWriter(path, false, System.Text.Encoding.Default);
					sz.Serialize(fs, ss);
					fs.Close();
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return (ret);
			}
		}

		public void UpdateData(string InfoStr)
		{
			string[] cols = InfoStr.Split(',');

			//ListViewに追加
			operatorName = cols[0];
			seikeikiNo = cols[1];//成型機
			goukiNo = cols[2];//成型機No
			hinshu = cols[3];//製品品種
			sleeveNo = cols[4];//スリーブNo
			cavNo = cols[5];//CavNo
			priority = cols[6];//優先度

			DateTime dt = DateTime.Now;
			dateInfo = dt.ToString("yyyy/MM/dd");
			timeInfo = dt.ToString("HH:mm:ss");
			tairyuTime = "";
			status = "測定待ち";
			endDate = "";
			endTime = "";
			
			string[] item1 = {dateInfo, timeInfo, operatorName, seikeikiNo, goukiNo, hinshu, sleeveNo, cavNo, tairyuTime, status, endDate, endTime, priority};

			listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加

			if(priority == "通常")
			{
				listView1.Items[0].BackColor = Color.Lime;//背景色
				colorList.Insert(0, Color.Lime);
			}
			else if(priority == "最優先")
			{
				listView1.Items[0].BackColor = Color.Red;//背景色
				colorList.Insert(0, Color.Red);
			}
			else if(priority == "優先")
			{
				listView1.Items[0].BackColor = Color.Yellow;//背景色
				colorList.Insert(0, Color.Yellow);
			}
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

			int openCount = 0;
			int greenCount = 0;
			int orangeCount = 0;
			int magentaCount = 0;
			int redCount = 0;
			int yellowCount = 0;
			for(int i = 0; i < listView1.Items.Count; i++)
			{
				if(listView1.Items[i].SubItems[9].Text == "測定待ち")
				{
					openCount++;
				}
				if(listView1.Items[i].BackColor == Color.Lime)
				{
					greenCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Orange)
				{
					orangeCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Magenta)
				{
					magentaCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Red)
				{
					redCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Yellow)
				{
					yellowCount++;
				}
			}
			label5.Text = string.Format("{0}", openCount);

			label6.Text = string.Format("{0}", greenCount);
			label7.Text = string.Format("{0}", orangeCount);
			label8.Text = string.Format("{0}", magentaCount);
			label11.Text = string.Format("{0}", redCount);
			label13.Text = string.Format("{0}", yellowCount);

			if(redCount == 0)
			{
				timer5.Enabled = false;
				label10.BackColor = Color.Red;
				label11.BackColor = Color.Red;
				label9.BackColor = Color.White;
			}
			else
			{
				timer5.Enabled = true;
				label9.BackColor = Color.Red;
			}

			if(yellowCount == 0)
			{
				timer6.Enabled = false;
				label12.BackColor = Color.Yellow;
				label13.BackColor = Color.Yellow;
				label9.ForeColor = Color.Black;
			}
			else
			{
				timer6.Enabled = true;
				label9.ForeColor = Color.Yellow;
			}

			//CSVに出力する
			string logBuf = dateInfo + "," + timeInfo + "," + operatorName + "," + seikeikiNo + "," + goukiNo + "," + hinshu + "," + sleeveNo + "," + cavNo + "," + tairyuTime + "," + status + "," + endDate + "," + endTime + "," + priority;
			WriteDataToCsv(logBuf);

			//ログに出力する
			StatusFileOut(string.Format("{0}", openCount));

			//追加登録可否
			DialogResult result = MessageBox.Show("追加登録しますか？", "確認♪", MessageBoxButtons.OKCancel);
			if(result == DialogResult.OK)
			{
				timer4.Enabled = true;
			}
			else
			{
				timer1.Enabled = true;
			}
		}

        private void button1_Click(object sender, EventArgs e)
        {
			timer1.Enabled = false;
			//登録画面の表示
            Form2 form2 = new Form2();
            form2.ShowDialog();
            
            //登録画面より情報を取得
            InfoStr = form2.EditInfo;
			form2.Dispose();
            if(InfoStr == "")
            {
				timer1.Enabled = true;
				return;
			}
			UpdateData(InfoStr);
		}

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
		    if(e.Button == System.Windows.Forms.MouseButtons.Right)
			{
				return;
			}

			timer1.Enabled = false;
			int index = 0;
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
			    index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

				operatorName = listView1.Items[index].SubItems[2].Text;//成型機
				seikeikiNo = listView1.Items[index].SubItems[3].Text;//成型機
				goukiNo = listView1.Items[index].SubItems[4].Text;//成型機No
				hinshu = listView1.Items[index].SubItems[5].Text;//製品品種
				sleeveNo = listView1.Items[index].SubItems[6].Text;//スリーブNo
				cavNo = listView1.Items[index].SubItems[7].Text;//CavNo
                status = listView1.Items[index].SubItems[9].Text;//状態

				if(listView1.Items[index].BackColor == Color.Red)
				{
					priority = "最優先";
				}
				else if(listView1.Items[index].BackColor == Color.Yellow)
				{
					priority = "優先";
				}
				else
				{
					priority = "通常";
				}


                if(status == "測定終了")
                {
					timer1.Enabled = true;
                    return;
                }

                //登録・編集画面の表示
                Form2 form2 = new Form2();
				form2.SetInfo(operatorName, seikeikiNo, goukiNo, hinshu, sleeveNo, cavNo, priority);
	            form2.ShowDialog();

	            //登録・編集画面より情報を取得
	            InfoStr = form2.EditInfo;
				form2.Dispose();
	            if(InfoStr == "")
	            {
					timer1.Enabled = true;
					return;
				}

				string[] cols = InfoStr.Split(',');

				//ListViewに追加
				operatorName = cols[0];//登録者
				seikeikiNo = cols[1];//成型機
				goukiNo = cols[2];//成型機No
				hinshu = cols[3];//製品品種
				sleeveNo = cols[4];//スリーブNo
				cavNo = cols[5];//CavNo
				priority = cols[6];//優先度

				//CSVの更新
                StreamReader reader = null;
                StreamWriter writer = null;
                string line = "";
                string path = "";
				try
				{
	                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                    string[] lines = File.ReadAllLines(currentCsvFile);
                    int lineMax = lines.Length;//CSVの行数取得

                    path = currentCsvFile + ".tmp";
					writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

					int count = 0;
					while(reader.Peek() >= 0)
					{
						line = reader.ReadLine();

						if((lineMax - 1) - index == count)
						{
							string[] linesub = line.Split(',');
							string buf = "";
							for(int i = 0; i < linesub.Length; i++)
							{
								if(i == 0)
								{
									buf = linesub[i];
								}
								else if(i == 2)//登録者
								{
									buf += "," + operatorName;
								}
								else if(i == 3)//成型機
								{
									buf += "," + seikeikiNo;
								}
								else if(i == 4)//号機No
								{
									buf += "," + goukiNo;
								}
								else if(i == 5)//製品品種
								{
                                    buf += "," + hinshu;
								}
								else if(i == 6)//スリーブNo
								{
									buf += "," + sleeveNo;
								}
								else if(i == 7)//CavNo
								{
									buf += "," + cavNo;
								}
								else if(i == 12)//優先度
								{
									buf += "," + priority;
								}
								else
								{
									buf += "," + linesub[i];
								}
							}
							writer.WriteLine(buf);
						}
						else
						{
							writer.WriteLine(line);
						}

						count++;
					}

					listView1.Items[index].SubItems[2].Text = operatorName;
					listView1.Items[index].SubItems[3].Text = seikeikiNo;
					listView1.Items[index].SubItems[4].Text = goukiNo;
					listView1.Items[index].SubItems[5].Text = hinshu;
					listView1.Items[index].SubItems[6].Text = sleeveNo;
					listView1.Items[index].SubItems[7].Text = cavNo;

					if(priority == "通常")//滞留時間に沿った色に戻す
					{
						string strTime = listView1.Items[index].SubItems[0].Text + " " + listView1.Items[index].SubItems[1].Text;
						DateTime dTime = DateTime.Parse(strTime);
						DateTime dt = DateTime.Now;
						TimeSpan ts = dt - dTime;

                        if(ts.Days < SETDATA.dayLimit)
                        {
							if(0 < ts.Days)
							{
								listView1.Items[index].BackColor = Color.Orange;//背景色
								colorList[index] = Color.Orange;
							}
							else
							{
								if(ts.Hours < SETDATA.hourLimit)
								{
									listView1.Items[index].BackColor = Color.Lime;//背景色
									colorList[index] = Color.Lime;
								}
		                        else
								{
									listView1.Items[index].BackColor = Color.Orange;//背景色
									colorList[index] = Color.Orange;
								}
							}
						}
						else
						{
							listView1.Items[index].BackColor = Color.Magenta;//背景色
							colorList[index] = Color.Magenta;
						}
					}
					else if(priority == "最優先")
					{
						listView1.Items[index].BackColor = Color.Red;//背景色
						colorList[index] = Color.Red;
					}
					else if(priority == "優先")
					{
						listView1.Items[index].BackColor = Color.Yellow;//背景色
						colorList[index] = Color.Yellow;
					}

		            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
				}
				catch (System.IO.IOException ex)
				{
					string errorStr = "CSVファイルを開けなかった可能性があります";
				    System.Console.WriteLine(errorStr);
			        System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
				}
				finally
				{
					if(reader != null)
					{
						reader.Close();
						//元ファイル削除
						File.Delete(@currentCsvFile);
					}
					if(writer != null)
					{
						writer.Close();
						//一時ファイル→元ファイルへファイル名変更
						System.IO.File.Move(@path, @currentCsvFile);
					}

				}


			}

			int openCount = 0;
			int greenCount = 0;
			int orangeCount = 0;
			int magentaCount = 0;
			int redCount = 0;
			int yellowCount = 0;
			for(int i = 0; i < listView1.Items.Count; i++)
			{
				if(listView1.Items[i].SubItems[9].Text == "測定待ち")
				{
					openCount++;
				}
				if(listView1.Items[i].BackColor == Color.Lime)
				{
					greenCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Orange)
				{
					orangeCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Magenta)
				{
					magentaCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Red)
				{
					redCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Yellow)
				{
					yellowCount++;
				}
			}
			label5.Text = string.Format("{0}", openCount);

			label6.Text = string.Format("{0}", greenCount);
			label7.Text = string.Format("{0}", orangeCount);
			label8.Text = string.Format("{0}", magentaCount);
			label11.Text = string.Format("{0}", redCount);
			label13.Text = string.Format("{0}", yellowCount);

			if(redCount == 0)
			{
				timer5.Enabled = false;
				label10.BackColor = Color.Red;
				label11.BackColor = Color.Red;
				label9.BackColor = Color.White;
			}
			else
			{
				timer5.Enabled = true;
				label9.BackColor = Color.Red;
			}

			if(yellowCount == 0)
			{
				timer6.Enabled = false;
				label12.BackColor = Color.Yellow;
				label13.BackColor = Color.Yellow;
				label9.ForeColor = Color.Black;
			}
			else
			{
				timer6.Enabled = true;
				label9.ForeColor = Color.Yellow;
			}

			timer1.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
			timer1.Enabled = false;
			int index = 0;
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
			    index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

				operatorName = listView1.Items[index].SubItems[2].Text;
				seikeikiNo = listView1.Items[index].SubItems[3].Text;
				goukiNo = listView1.Items[index].SubItems[4].Text;
				hinshu = listView1.Items[index].SubItems[5].Text;
				sleeveNo = listView1.Items[index].SubItems[6].Text;
				cavNo = listView1.Items[index].SubItems[7].Text;
                status = listView1.Items[index].SubItems[9].Text;//状態

                if(status == "測定終了")
                {
					timer1.Enabled = true;
                    return;
                }

				string delColumn = "測定終了したスリーブは" + "\r\n";
	            delColumn += string.Format("登録者　：　　{0}さん", operatorName) + "\r\n";
	            delColumn += string.Format("成型機　：　　{0}{1}号機", seikeikiNo, goukiNo) + "\r\n";
				delColumn += string.Format("品種　：　{0}", hinshu) + "\r\n";
	            delColumn += string.Format("スリーブ　：　{0}", sleeveNo) + "\r\n";
	            delColumn += string.Format("Cav　：　{0}", cavNo) + "\r\n";
	            delColumn += "で間違いありませんか？";
				
	            DialogResult result = MessageBox.Show(delColumn, "測定終了　選択", MessageBoxButtons.YesNo);

				if(result == DialogResult.Yes)
				{
					//CSVの更新
	                StreamReader reader = null;
	                StreamWriter writer = null;
	                string line = "";
	                string path = "";
	                string endStatus = "測定終了";
					DateTime dt = DateTime.Now;
					dateInfo = dt.ToString("yyyy/MM/dd");
					timeInfo = dt.ToString("HH:mm:ss");
					try
					{
		                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
	                    string[] lines = File.ReadAllLines(currentCsvFile);
	                    int lineMax = lines.Length;//CSVの行数取得

	                    path = currentCsvFile + ".tmp";
						writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

						int count = 0;
						int csvLen = 0;
						while(reader.Peek() >= 0)
						{
							line = reader.ReadLine();

							if(count == 0)
							{
								csvLen = line.Split(',').Length;
							}

							if((lineMax - 1) - index == count)
							{
								string[] linesub = line.Split(',');
								string buf = "";

								for(int i = 0; i < csvLen; i++)
								{
									if(i == 0)
									{
										buf = linesub[i];
									}
									else if(i == 9)//状態
									{
										buf += "," + endStatus;
									}
									else if(i == 10)//終了日
									{
										buf += "," + dateInfo;
									}
									else if(i == 11)//終了時間
									{
										buf += "," + timeInfo;
									}
									else
									{
										buf += "," + linesub[i];
									}
								}
								writer.WriteLine(buf);
							}
							else
							{
								writer.WriteLine(line);
							}

							count++;
						}

			            listView1.Items[index].ForeColor = Color.Black;//初期の色
			            listView1.Items[index].BackColor = Color.Gray;//背景色

						listView1.Items[index].SubItems[9].Text = endStatus;
						listView1.Items[index].SubItems[10].Text = dateInfo;
						listView1.Items[index].SubItems[11].Text = timeInfo;
			            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
			            
			            colorList[index] = Color.Gray;
					}
					catch (System.IO.IOException ex)
					{
						string errorStr = "CSVファイルを開けなかった可能性があります";
					    System.Console.WriteLine(errorStr);
				        System.Console.WriteLine(ex.Message);
						LogFileOut(errorStr);
					}
					finally
					{
						if(reader != null)
						{
							reader.Close();
							//元ファイル削除
							File.Delete(@currentCsvFile);
						}
						if(writer != null)
						{
							writer.Close();
							//一時ファイル→元ファイルへファイル名変更
							System.IO.File.Move(@path, @currentCsvFile);
						}

					}
				}
			}
			else
			{
				MessageBox.Show("測定終了したスリーブを選択して下さい");
				timer1.Enabled = true;
				return;
			}

			int openCount = 0;
			int greenCount = 0;
			int orangeCount = 0;
			int magentaCount = 0;
			int redCount = 0;
			int yellowCount = 0;
			for(int i = 0; i < listView1.Items.Count; i++)
			{
				if(listView1.Items[i].SubItems[9].Text == "測定待ち")
				{
					openCount++;
				}
				if(listView1.Items[i].BackColor == Color.Lime)
				{
					greenCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Orange)
				{
					orangeCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Magenta)
				{
					magentaCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Red)
				{
					redCount++;
				}
				else if(listView1.Items[i].BackColor == Color.Yellow)
				{
					yellowCount++;
				}
			}
			label5.Text = string.Format("{0}", openCount);

			label6.Text = string.Format("{0}", greenCount);
			label7.Text = string.Format("{0}", orangeCount);
			label8.Text = string.Format("{0}", magentaCount);
			label11.Text = string.Format("{0}", redCount);
			label13.Text = string.Format("{0}", yellowCount);

			//ログに出力する
			StatusFileOut(string.Format("{0}", openCount));

			int redTotal = 0;
			int yellowTotal = 0;
			for(int i = 0; i < colorList.Count; i++)
			{
				if(colorList[i] == Color.Red)
				{
					redTotal++;
				}
				else if(colorList[i] == Color.Yellow)
				{
					yellowTotal++;
				}
			}
			if(redTotal == 0)
			{
				timer5.Enabled = false;
				label10.BackColor = Color.Red;
				label11.BackColor = Color.Red;
				label9.BackColor = Color.White;
			}
			if(yellowTotal == 0)
			{
				timer6.Enabled = false;
				label12.BackColor = Color.Yellow;
				label13.BackColor = Color.Yellow;
				label9.ForeColor = Color.Black;
			}

			timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
			//CSVの更新
            StreamWriter writer = null;
            string path = "";
			int openCount = 0;
			int greenCount = 0;
			int orangeCount = 0;
			int magentaCount = 0;
			int redCount = 0;
			int yellowCount = 0;

			//ListViewを一度クリアする
			listView1.Items.Clear();
			
//			colorList.Clear();

			string[] item1 = {dateInfo, timeInfo, operatorName, seikeikiNo, goukiNo, hinshu, sleeveNo, cavNo, tairyuTime, status, endDate, endTime, priority};

            //CSVファイルのバックアップを行う。今のCSVを日付のCSVで作成し、測定待ちだけ残して再作成
            DateTime d = DateTime.Now;
            if(!isUpdate)
			{
				if(targetDate <= d)
				{
					isUpdate = true;
					targetDate = d.AddDays(1);
				}
			}
			else
			{
				isUpdate = false;
				timer1.Interval = SETDATA.timerOut * 1000;
			}

            try
            {
	            //CSV読込。高速化を狙い、最大行数を取得後、for分でループする
	            var readToEnd = File.ReadAllLines(@currentCsvFile, Encoding.GetEncoding("Shift_JIS"));
	            int lines = readToEnd.Length;

                string[] strlines = File.ReadAllLines(currentCsvFile);
                int lineMax = strlines.Length;//CSVの行数取得

                path = currentCsvFile + ".tmp";
				writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

	            for (int i = 0; i < lines; i++)
	            {
	                //１行のstringをstream化してTextFieldParserで処理する
	                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
	                {
                        using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
                        {
                            parser.TextFieldType = FieldType.Delimited;
                            parser.Delimiters = new[] { "," };
                            parser.HasFieldsEnclosedInQuotes = true;
                            parser.TrimWhiteSpace = false;
                            string[] fields = parser.ReadFields();
                            string buf = "";

                            if (i == 0)//ヘッダ部
                            {
                                for (int j = 0; j < fields.Length; j++)
                                {
                                    if (j == 0)
                                    {
                                        buf = fields[j];
                                    }
                                    else
                                    {
                                        buf += "," + fields[j];
                                    }

                                    item1[j] = fields[j];
                                }

                                writer.WriteLine(buf);
                            }
                            else
                            {

                                if(isUpdate)
                                {
                                    if (fields[9] == "測定終了")//状態
                                    {
										int tag = (lineMax - 1) - i;
										colorList.RemoveAt(tag);
										continue;
									}
								}
                                
                                string tairyu = "";
                                for (int j = 0; j < fields.Length; j++)
                                {
                                    if (j == 0)
                                    {
                                        buf = fields[j];
                                    }
                                    else if (j == 8)//滞留時間
                                    {
                                        if (fields[9] == "測定終了")//状態
                                        {
                                            buf += "," + fields[j];
                                            item1[j] = fields[j];
                                            continue;
                                        }

                                        string strTime = fields[0] + " " + fields[1];
                                        DateTime dTime = DateTime.Parse(strTime);
                                        DateTime dt = DateTime.Now;
                                        TimeSpan ts = dt - dTime;

                                        tairyu = string.Format("{0}日{1}時間{2}分{3}秒", ts.Days, ts.Hours, ts.Minutes, ts.Seconds);

                                        buf += "," + tairyu;
                                    }
                                    else
                                    {
                                        buf += "," + fields[j];
                                    }

                                    item1[j] = fields[j];
                                }
                                writer.WriteLine(buf);
                                listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
                            }

                            if (fields[9] == "測定待ち")
                            {
                                openCount++;
                            }
                        }
	                }

                    if (i == 0)//ヘッダ部
					{
						continue;
					}

					if(item1[9] == "測定終了")
					{
						listView1.Items[0].BackColor = Color.Gray;//背景色

						int tag = (lineMax - 1) - i;
						colorList[tag] = Color.Gray;
					}
					else
					{
						int tag = (lineMax - 1) - i;
						if(colorList[tag] == Color.Red)
						{
							listView1.Items[0].BackColor = Color.Red;
							redCount++;
							continue;
						}
						else if(colorList[tag] == Color.Yellow)
						{
							listView1.Items[0].BackColor = Color.Yellow;
							yellowCount++;
							continue;
						}

						string strTime = item1[0] + " " + item1[1];
						DateTime dTime = DateTime.Parse(strTime);
						DateTime dt = DateTime.Now;
						TimeSpan ts = dt - dTime;

                        if(ts.Days < SETDATA.dayLimit)
                        {
							if(0 < ts.Days)
							{
								listView1.Items[0].BackColor = Color.Orange;//背景色

								tag = (lineMax - 1) - i;
								colorList[tag] = Color.Orange;

								orangeCount++;
							}
							else
							{
								if(ts.Hours < SETDATA.hourLimit)
								{
									listView1.Items[0].BackColor = Color.Lime;//背景色

									tag = (lineMax - 1) - i;
									colorList[tag] = Color.Lime;

									greenCount++;
								}
		                        else
								{
									listView1.Items[0].BackColor = Color.Orange;//背景色

									tag = (lineMax - 1) - i;
									colorList[tag] = Color.Orange;

									orangeCount++;
								}
							}
						}
						else
						{
							listView1.Items[0].BackColor = Color.Magenta;//背景色

							tag = (lineMax - 1) - i;
							colorList[tag] = Color.Magenta;

							magentaCount++;
						}
					}
	            }

				label5.Text = string.Format("{0}", openCount);
				label6.Text = string.Format("{0}", greenCount);
				label7.Text = string.Format("{0}", orangeCount);
				label8.Text = string.Format("{0}", magentaCount);
				label11.Text = string.Format("{0}", redCount);
				label13.Text = string.Format("{0}", yellowCount);
			}
			catch (System.IO.IOException ex)
		    {
		        // ファイルを開くのに失敗したときエラーメッセージを表示
				string errorStr = "状変時にCSVファイルを開けなかった可能性があります";
			    System.Console.WriteLine(errorStr);
		        System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
			}
			finally
			{
				if(isUpdate)
				{
					//→元ファイルを日付ファイルへファイル名変更(バックアップ)
					string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
					string backUpDir = stCurrentDir + "\\backup";
					string result = d.ToString("yyyyMMdd_HHmmss");
					string backupCsv = backUpDir + "\\PgmMeasureInfo_" + result + ".csv";

					if(!Directory.Exists(backUpDir))
					{
						Directory.CreateDirectory(backUpDir);
					}

					//存在していなければ元ファイルをコピー
					if(!System.IO.File.Exists(backupCsv))
					{
						System.IO.File.Move(@currentCsvFile, @backupCsv);
					}

				}
				else
				{
					if(System.IO.File.Exists(currentCsvFile))
	                {
						//元ファイル削除
						File.Delete(@currentCsvFile);
					}
				}

				if(writer != null)
				{
					writer.Close();
					//一時ファイル→元ファイルへファイル名変更
					System.IO.File.Move(@path, @currentCsvFile);
				}

			}
			listView1.Font = new System.Drawing.Font("Times New Roman", 18, System.Drawing.FontStyle.Regular);
            //ヘッダの幅を自動調節
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

        }

        private void timer2_Tick(object sender, EventArgs e)//PCをスリープさせない用
        {
            //画面暗転阻止
            SetThreadExecutionState(ExecutionState.DisplayRequired);

            // ドラッグ操作の準備 (struct 配列の宣言)
            INPUT[] input = new INPUT[1];  // イベントを格納

            // ドラッグ操作の準備 (イベントの定義 = 相対座標へ移動)
            input[0].mi.dx = 0;  // 相対座標で0　つまり動かさない
            input[0].mi.dy = 0;  // 相対座標で0 つまり動かさない
            input[0].mi.dwFlags = MOUSEEVENTF_MOVED;

            // ドラッグ操作の実行 (イベントの生成)
            SendInput(1, input, Marshal.SizeOf(input[0]));

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
			DialogResult result = MessageBox.Show("閉じてよいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			if(result == DialogResult.No)
			{
				e.Cancel = true;
			}
		}

        private void timer3_Tick(object sender, EventArgs e)//現在時刻更新
        {
            DateTime d = DateTime.Now;
            label9.Text = string.Format("{0}/{1}/{2}({3})", d.Year, d.Month, d.Day, d.ToString("ddd")) + "\r\n" + string.Format("{0:00}:{1:00}:{2:00}", d.Hour, d.Minute, d.Second);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DateTime d = DateTime.Now;
			targetDate = d.AddDays(-1);

			timer1.Enabled = false;
			timer1.Interval = 500;
			timer1.Enabled = true;
        }

        private void timer4_Tick(object sender, EventArgs e)//追加登録時
        {
			timer4.Enabled = false;
            //登録画面の表示
            Form2 form2 = new Form2();
            cavNo = "";//追加登録時はcavはクリア
			form2.SetInfo(operatorName, seikeikiNo, goukiNo, hinshu, sleeveNo, cavNo, priority);
            form2.ShowDialog();

            InfoStr = form2.EditInfo;
			form2.Dispose();
            if(InfoStr == "")
            {
				timer1.Enabled = true;
				return;
			}

			UpdateData(InfoStr);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)//優先を登録
        {
			UpdatePriorityStatus("優先");
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
				int index = listView1.SelectedItems[0].Index;
                listView1.Items[index].BackColor = Color.Yellow;

				colorList[index] = Color.Yellow;

				//背景色をカウント
				int greenCount = 0;
				int orangeCount = 0;
				int magentaCount = 0;
				int redCount = 0;
				int yellowCount = 0;

				for(int i = 0; i < listView1.Items.Count; i++)
				{
					if(listView1.Items[i].BackColor == Color.Lime)
					{
						greenCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Orange)
					{
						orangeCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Magenta)
					{
						magentaCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Red)
					{
						redCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Yellow)
					{
						yellowCount++;
					}
				}

				label6.Text = string.Format("{0}", greenCount);
				label7.Text = string.Format("{0}", orangeCount);
				label8.Text = string.Format("{0}", magentaCount);
				label11.Text = string.Format("{0}", redCount);
				label13.Text = string.Format("{0}", yellowCount);

				if(redCount == 0)
				{
					timer5.Enabled = false;
					label10.BackColor = Color.Red;
					label11.BackColor = Color.Red;
					label9.BackColor = Color.White;
				}
				else
				{
					timer5.Enabled = true;
					label9.BackColor = Color.Red;
				}

				if(yellowCount == 0)
				{
					timer6.Enabled = false;
					label12.BackColor = Color.Yellow;
					label13.BackColor = Color.Yellow;
					label9.ForeColor = Color.Black;
				}
				else
				{
					timer6.Enabled = true;
					label9.ForeColor = Color.Yellow;
				}
			}
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)//解除
        {
			UpdatePriorityStatus("通常");
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
				int index = listView1.SelectedItems[0].Index;

				string strTime = listView1.Items[index].SubItems[0].Text + " " + listView1.Items[index].SubItems[1].Text;
				DateTime dTime = DateTime.Parse(strTime);
				DateTime dt = DateTime.Now;
				TimeSpan ts = dt - dTime;

                if(ts.Days < SETDATA.dayLimit)
                {
					if(0 < ts.Days)
					{
						listView1.Items[index].BackColor = Color.Orange;//背景色
						colorList[index] = Color.Orange;
					}
					else
					{
						if(ts.Hours < SETDATA.hourLimit)
						{
							listView1.Items[index].BackColor = Color.Lime;//背景色
							colorList[index] = Color.Lime;
						}
                        else
						{
							listView1.Items[index].BackColor = Color.Orange;//背景色
							colorList[index] = Color.Orange;
						}
					}
				}
				else
				{
					listView1.Items[index].BackColor = Color.Magenta;//背景色
					colorList[index] = Color.Magenta;
				}

				//背景色をカウント
				int greenCount = 0;
				int orangeCount = 0;
				int magentaCount = 0;
				int redCount = 0;
				int yellowCount = 0;
				for(int i = 0; i < listView1.Items.Count; i++)
				{
					if(listView1.Items[i].BackColor == Color.Lime)
					{
						greenCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Orange)
					{
						orangeCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Magenta)
					{
						magentaCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Red)
					{
						redCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Yellow)
					{
						yellowCount++;
					}
				}

				label6.Text = string.Format("{0}", greenCount);
				label7.Text = string.Format("{0}", orangeCount);
				label8.Text = string.Format("{0}", magentaCount);
				label11.Text = string.Format("{0}", redCount);
				label13.Text = string.Format("{0}", yellowCount);

				if(redCount == 0)
				{
					timer5.Enabled = false;
					label10.BackColor = Color.Red;
					label11.BackColor = Color.Red;
					label9.BackColor = Color.White;
				}
				else
				{
					timer5.Enabled = true;
					label9.BackColor = Color.Red;
				}

				if(yellowCount == 0)
				{
					timer6.Enabled = false;
					label12.BackColor = Color.Yellow;
					label13.BackColor = Color.Yellow;
					label9.ForeColor = Color.Black;
				}
				else
				{
					timer6.Enabled = true;
					label9.ForeColor = Color.Yellow;
				}
			}
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)//最優先を登録
        {
			UpdatePriorityStatus("最優先");
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
				int index = listView1.SelectedItems[0].Index;
                listView1.Items[index].BackColor = Color.Red;

				colorList[index] = Color.Red;

				//背景色をカウント
				int greenCount = 0;
				int orangeCount = 0;
				int magentaCount = 0;
				int redCount = 0;
				int yellowCount = 0;
				for(int i = 0; i < listView1.Items.Count; i++)
				{
					if(listView1.Items[i].BackColor == Color.Lime)
					{
						greenCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Orange)
					{
						orangeCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Magenta)
					{
						magentaCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Red)
					{
						redCount++;
					}
					else if(listView1.Items[i].BackColor == Color.Yellow)
					{
						yellowCount++;
					}
				}

				label6.Text = string.Format("{0}", greenCount);
				label7.Text = string.Format("{0}", orangeCount);
				label8.Text = string.Format("{0}", magentaCount);
				label11.Text = string.Format("{0}", redCount);
				label13.Text = string.Format("{0}", yellowCount);

				if(redCount == 0)
				{
					timer5.Enabled = false;
					label10.BackColor = Color.Red;
					label11.BackColor = Color.Red;
					label9.BackColor = Color.White;
				}
				else
				{
					timer5.Enabled = true;
					label9.BackColor = Color.Red;
				}

				if(yellowCount == 0)
				{
					timer6.Enabled = false;
					label12.BackColor = Color.Yellow;
					label13.BackColor = Color.Yellow;
					label9.ForeColor = Color.Black;
				}
				else
				{
					timer6.Enabled = true;
					label9.ForeColor = Color.Yellow;
				}
			}
        }

		private void UpdatePriorityStatus(string priorityStatus)
		{
			timer1.Enabled = false;
			int index = 0;
			if(listView1.SelectedItems.Count > 0)//ListViewに1つでも登録がある
			{
			    index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

				//CSVの更新
                StreamReader reader = null;
                StreamWriter writer = null;
                string line = "";
                string path = "";
				try
				{
	                reader = new StreamReader(currentCsvFile, System.Text.Encoding.GetEncoding("Shift_JIS"));
                    string[] lines = File.ReadAllLines(currentCsvFile);
                    int lineMax = lines.Length;//CSVの行数取得

                    path = currentCsvFile + ".tmp";
					writer = new StreamWriter(path, false, System.Text.Encoding.GetEncoding("Shift_JIS"));

					int count = 0;
					while(reader.Peek() >= 0)
					{
						line = reader.ReadLine();

						if((lineMax - 1) - index == count)
						{
							string[] linesub = line.Split(',');
							string buf = "";
							for(int i = 0; i < linesub.Length; i++)
							{
								if(i == 0)
								{
									buf = linesub[i];
								}
								else if(i == 12)//優先度
								{
									buf += "," + priorityStatus;
								}
								else
								{
									buf += "," + linesub[i];
								}
							}
							writer.WriteLine(buf);
						}
						else
						{
							writer.WriteLine(line);
						}
						count++;
					}
				}
				catch (System.IO.IOException ex)
				{
					string errorStr = "CSVファイルを開けなかった可能性があります";
				    System.Console.WriteLine(errorStr);
			        System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
				}
				finally
				{
					if(reader != null)
					{
						reader.Close();
						//元ファイル削除
						File.Delete(@currentCsvFile);
					}
					if(writer != null)
					{
						writer.Close();
						//一時ファイル→元ファイルへファイル名変更
						System.IO.File.Move(@path, @currentCsvFile);
					}

				}
			}

			timer1.Enabled = true;
		}

        private void listView1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                if (listView1.SelectedItems.Count > 0)
                {
					int index = listView1.SelectedItems[0].Index;
	                string status = listView1.Items[index].SubItems[9].Text;//状態

	                if(status == "測定終了")
	                {
	                    return;
	                }

                    System.Drawing.Point p = System.Windows.Forms.Cursor.Position;
                    this.contextMenuStrip1.Show(p);
                }
            }
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
			Color color = label10.BackColor;
			color = Color.FromArgb(color.ToArgb() ^ 0xFFFFFF);//反転
			label10.BackColor = color;
			label11.BackColor = color;

			Color watch_color = label9.BackColor;
			watch_color = Color.FromArgb(watch_color.ToArgb() ^ 0xFFFFFF);//反転
			label9.BackColor = watch_color;
        }

        private void timer6_Tick(object sender, EventArgs e)
        {
			Color color = label12.BackColor;
			color = Color.FromArgb(color.ToArgb() ^ 0xFFFFFF);//反転
			label12.BackColor = color;
			label13.BackColor = color;

			Color watch_color = label9.BackColor;
			watch_color = Color.FromArgb(watch_color.ToArgb() ^ 0xFFFFFF);//反転
			label9.BackColor = watch_color;
        }
    }
}
