using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Text.RegularExpressions;

namespace YanBo
{
    class FormService
    {
        public static int scheme = 0;
        public String SaveFile(String title, String filter)
        {
            String path = "";
            using (Form1 form = new Form1())
            {
                form.saveFileDialog1.Title = title;
                form.saveFileDialog1.Filter = filter;
                DialogResult dr = form.saveFileDialog1.ShowDialog();

                form.saveFileDialog1.RestoreDirectory = true;
                if (dr == DialogResult.OK)
                {
                    path = form.saveFileDialog1.FileName;
                }

                if (path.Equals(""))
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("û��ѡ��Ҫ���ɵ��ļ�");
                }
            }
            return path;
        }
        public String OpenFile(String title, String filter)
        {
            String path = "";
            using (Form1 form = new Form1())
            {
                form.openFileDialog1.Title = title;
                form.openFileDialog1.Filter = filter;

                DialogResult dr = form.openFileDialog1.ShowDialog();

                form.openFileDialog1.RestoreDirectory = true;
                if (dr == DialogResult.OK)
                {
                    path = form.openFileDialog1.FileName;
                }

                if (path.Equals(""))
                {
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("û��ѡ��Ҫ�򿪵��ļ�");
                }
            }
            return path;
        }
        public String OpenDir(String title)
        {
            FolderBrowserDialog dir = new FolderBrowserDialog();
            if (dir.ShowDialog() == DialogResult.OK)
            {
                return dir.SelectedPath;
            }
            return "";
        }

        public Object[] ShowFourAddressForm()
        {
            FourAdressForm form = new FourAdressForm();
            bool flag = true;
            form.ShowDialog();
            Object[] formText = new Object[10];
            /**0���ļ���ַ
             * 1���Ƿ�Ϊ�������
             * 2���������
             * 3���Ƿ񸲸���ǰ����
            */
            if (!File.Exists(form.textBox1.Text))
            {
                MessageBox.Show("�ļ�������");
                flag = false;
            }
            formText[0] = form.textBox1.Text;
            bool surface = form.radioButton1.Checked;
            formText[1] = surface;
            double di = 0.03;
            if (!surface)
            {
                scheme = 1;
                String distance = form.textBox2.Text;
                string regexString = @"(^[0-9]*[1-9][0-9]*$)|(^([0-9]{1,}[.][0-9]*)$)";//д������ʽ��ֻ����������&С��
                Match m = Regex.Match(distance, regexString);

                if (m.Success == false)
                {

                    MessageBox.Show("ֻ����������������С��");
                    flag = false;
                }
                else
                {
                    di = double.Parse(distance);
                    if (di < 0)
                    {
                        MessageBox.Show("ֻ���������0������������С��");
                        flag = false;
                    }
                }

            }
            formText[2] = di;
            formText[3] = form.checkBox1.Checked;
            if (flag == false)
            {
                return null;
            }

            return formText;
        }
    }
}
