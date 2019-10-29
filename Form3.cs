using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MeasureStatusMonitor
{
    public partial class Form3 : Form
    {
		public string EditInfo = "";
        public Form3()
        {
            InitializeComponent();

			for(int i = 0; i < Form1.SETDATA.registerOperator.Length; i++)
			{
	            comboBox5.Items.Add(Form1.SETDATA.registerOperator[i]);
			}
        }

        private void button1_Click(object sender, EventArgs e)
        {
			if(comboBox5.SelectedIndex == -1)
			{
				MessageBox.Show("担当者を選択して下さい");
				return;
			}
			if(comboBox1.SelectedIndex == -1)
			{
				MessageBox.Show("結果を選択して下さい");
				return;
			}

            string operatorName = string.Format("入力者　：　{0}さん", comboBox5.Text);
            operatorName += "\r\n";

            string ok_ng = string.Format("結果　：　{0}", comboBox1.Text);
            ok_ng += "\r\n";

            string combine = operatorName + ok_ng + "で間違いありませんか？";

            DialogResult result = MessageBox.Show(combine, "削除入力　削除", MessageBoxButtons.YesNo);
			if(result == DialogResult.Yes)
			{   
				EditInfo = comboBox5.Text + "," + comboBox1.Text;
                this.Close();
	        }


        }

		public void SetInfo(string operatorName, string seikeikiNo, string goukiNo, string hinshu, string sleeveNo, string cavNo)
		{
			this.Text = "担当者と結果を入力して下さい";
			
			label1.Text = string.Format("登録者　　：　{0}さん", operatorName);
			label2.Text = string.Format("成型機　　：　{0}{1}号機", seikeikiNo, goukiNo);
			label3.Text = string.Format("品種　　　：　{0}", hinshu);
			label4.Text = string.Format("スリーブ　：　{0}", sleeveNo);
			label5.Text = string.Format("Cav　　 　：　{0}", cavNo);
		}

    }
}
