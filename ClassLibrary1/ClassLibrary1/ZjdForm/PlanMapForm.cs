using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace YanBo.ZjdForm
{
    public partial class PlanMapForm : Form
    {
        public PlanMapForm()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PlanMapService planMapService = new PlanMapService();
            String berforStr = GetTextBoxString(textBox1);
            String afterStr = GetTextBoxString(textBox2);
            String dir = GetTextBoxString(textBox3);
            if (dir.Equals(""))
            {
                MessageBox.Show("你想干什么，文件夹名字没有选择");
            }
            else if (!Directory.Exists(dir))
            {
                MessageBox.Show("你想干什么，文件夹都不在，还想保存");
            }
            else
            {
                this.Close();
                this.Visible = false;
                planMapService.SplitDrawingForm(berforStr, afterStr, dir);
               
            }
           
        }
        public String GetTextBoxString(TextBox textBox)
        {
            String text = textBox.Text;
            if (text == null || text.Trim().Equals(""))
            {
                return "";
            }
            else
            {
                return text;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FormService formService = new FormService();
            String text = formService.OpenDir("选择要保存的位置");
            textBox3.Text = text;
        }
    }
}