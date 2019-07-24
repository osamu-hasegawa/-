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
    public partial class Form2 : Form
    {
		public string EditInfo = "";
        public Form2()
        {
            InitializeComponent();

			for(int i = 0; i < Form1.SETDATA.machineKind.Length; i++)
			{
	            comboBox1.Items.Add(Form1.SETDATA.machineKind[i]);
			}
			for(int i = 0; i < Form1.SETDATA.seihinName.Length; i++)
			{
	            comboBox3.Items.Add(Form1.SETDATA.seihinName[i]);
			}
			for(int i = 0; i < Form1.SETDATA.registerOperator.Length; i++)
			{
	            comboBox5.Items.Add(Form1.SETDATA.registerOperator[i]);
			}
			
			comboBox4.Text = "通常";
        }

        private void button1_Click(object sender, EventArgs e)
        {
			if(comboBox1.SelectedIndex == -1)
			{
				MessageBox.Show("成型機を選択して下さい");
				return;
			}
			if(comboBox3.SelectedIndex == -1)
			{
				MessageBox.Show("品種を選択して下さい");
				return;
			}
			if(comboBox5.SelectedIndex == -1)
			{
				MessageBox.Show("登録者を選択して下さい");
				return;
			}

            string operatorName = string.Format("登録者　：　{0}さん", comboBox5.Text);
            operatorName += "\r\n";

            string seikeiki = string.Format("成型機　：　{0}{1}号機", comboBox1.Text, numericUpDown1.Text);
            seikeiki += "\r\n";
            
            string hinshu = string.Format("品種　：　{0}", comboBox3.Text);
            hinshu += "\r\n";

            string sleeve = string.Format("スリーブ　：　{0}{1}", comboBox2.Text, numericUpDown2.Text);
            sleeve += "\r\n";

			string cavStr = "";
			int cavCount = 0;
			CheckBox [] checkTotal = {checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6};
			for(int i = 0; i < checkTotal.Length; i++)
			{
				if(checkTotal[i].Checked)
				{
					if(cavCount == 0)
					{
						cavStr += (i + 1);
					}
					else
					{
						cavStr += "/";
						cavStr += (i + 1);
					}
					cavCount++;
				}
			}
            string cav = string.Format("Cav　：　{0}", cavStr);
            cav += "\r\n";

            string priority = string.Format("優先度　：　{0}", comboBox4.Text);
            priority += "\r\n";

            string combine = operatorName + seikeiki + hinshu + sleeve + cav + priority + "で間違いありませんか？";

            DialogResult result = MessageBox.Show(combine, "測定待ちスリーブ　登録・編集", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{   
				EditInfo = comboBox5.Text + "," + comboBox1.Text + "," + numericUpDown1.Text + "," + comboBox3.Text + "," + comboBox2.Text + numericUpDown2.Text + "," + cavStr + "," + comboBox4.Text;
                this.Close();
	        }
        }

		public void SetInfo(string operatorName, string seikeikiNo, string goukiNo, string hinshu, string sleeveNo, string cavNo, string priority)
		{
			this.Text = "編集して下さい";
			comboBox1.Text = seikeikiNo;
			numericUpDown1.Text = goukiNo;

			comboBox5.Text = operatorName;
			comboBox3.Text = hinshu;
			comboBox4.Text = priority;

			if(sleeveNo.Length == 1)
			{
				numericUpDown2.Text = sleeveNo;
			}
			else if(sleeveNo.Length == 2)
			{
				comboBox2.Text = sleeveNo.Substring(0, 1);
				numericUpDown2.Text = sleeveNo.Substring(1, 1);
			}

			int [] cavMatrix = new int[]{0, 0, 0, 0, 0, 0};
			if(cavNo.Length == 0)
			{
				return;
			}
			else
			{
				if(cavNo.Length == 1)
				{
					cavMatrix[0] = int.Parse(cavNo);
				}
				else if(cavNo.Length == 3)
				{
					cavMatrix[0] = int.Parse(cavNo.Substring(0, 1));
					cavMatrix[1] = int.Parse(cavNo.Substring(2, 1));
				}
				else if(cavNo.Length == 5)
				{
					cavMatrix[0] = int.Parse(cavNo.Substring(0, 1));
					cavMatrix[1] = int.Parse(cavNo.Substring(2, 1));
					cavMatrix[2] = int.Parse(cavNo.Substring(4, 1));
				}
				else if(cavNo.Length == 7)
				{
					cavMatrix[0] = int.Parse(cavNo.Substring(0, 1));
					cavMatrix[1] = int.Parse(cavNo.Substring(2, 1));
					cavMatrix[2] = int.Parse(cavNo.Substring(4, 1));
					cavMatrix[3] = int.Parse(cavNo.Substring(6, 1));
				}
				else if(cavNo.Length == 9)
				{
					cavMatrix[0] = int.Parse(cavNo.Substring(0, 1));
					cavMatrix[1] = int.Parse(cavNo.Substring(2, 1));
					cavMatrix[2] = int.Parse(cavNo.Substring(4, 1));
					cavMatrix[3] = int.Parse(cavNo.Substring(6, 1));
					cavMatrix[4] = int.Parse(cavNo.Substring(8, 1));
				}
				else if(cavNo.Length == 11)
				{
					cavMatrix[0] = int.Parse(cavNo.Substring(0, 1));
					cavMatrix[1] = int.Parse(cavNo.Substring(2, 1));
					cavMatrix[2] = int.Parse(cavNo.Substring(4, 1));
					cavMatrix[3] = int.Parse(cavNo.Substring(6, 1));
					cavMatrix[4] = int.Parse(cavNo.Substring(8, 1));
					cavMatrix[5] = int.Parse(cavNo.Substring(10, 1));
				}

				CheckBox [] checkTotal = {checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6};
				for(int i = 0; i < cavMatrix.Length; i++)
				{
					if(cavMatrix[i] != 0)
					{
						checkTotal[cavMatrix[i] - 1].Checked = true;
					}
				}
	        }
		}


    }
}
