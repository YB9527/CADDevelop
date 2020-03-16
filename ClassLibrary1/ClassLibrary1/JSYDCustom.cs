using System;
using System.Collections.Generic;
using System.Text;
using YanBo_CG.ExcelUtils;
using NPOI.SS.UserModel;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace YanBo
{
    public class JSYDCustom
    {
        /// <summary>
        /// excel ת��Ϊ jsyd
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static IList<Jsyd> ReadExcel(string path)
        {

            IRow row;
            ICell cell;
            IList<Jsyd> list = new List<Jsyd>();
            if (!File.Exists(path))
            {
                return list;
            }
            ISheet sheet = ExcelRead.ReadExcelSheet(path, 0);
            for (int a = 4; a <= sheet.LastRowNum; a++)//��ȥһ�У��кϼ�
            {
                row = sheet.GetRow(a);

                string zdnum = row.GetCell(1).StringCellValue;
                if (zdnum != null && zdnum.Length == 19)
                {
                    Jsyd jsyd = new Jsyd();
                    jsyd.Zdnum = zdnum;
                    jsyd.QLRMC = GetStringValue(row.GetCell(3));
                    jsyd.Zdmj = GetDobleValue(row.GetCell(30));
                    jsyd.Syqmj = GetDobleValue(row.GetCell(31));
                    jsyd.Czmj = GetDobleValue(row.GetCell(34));
                    jsyd.JZZDZMJ = GetDobleValue(row.GetCell(34));
                    jsyd.TFH = GetStringValue(row.GetCell(28));
                    list.Add(jsyd);

                }
                else
                {
                    return list;
                }
                //����Ǻϼƣ�ֱ���˳�

            }
            return list;
        }

        private static double GetDobleValue(ICell cell)
        {
            if (cell == null)
            {
                return 0;
            }
            double num;
            cell.SetCellType(NPOI.SS.UserModel.CellType.String);
            if (double.TryParse(cell.StringCellValue, out num))
            {
                return Math.Round(num, 2);
            }
            return 0;
        }
        private static string GetStringValue(ICell cell)
        {
            if (cell == null)
            {
                return "";
            }

            cell.SetCellType(NPOI.SS.UserModel.CellType.String);
            return cell.StringCellValue;
        }

        internal static Dictionary<string, Jsyd> ListToDic(IList<Jsyd> jsyds)
        {
            Dictionary<string, Jsyd> dic = new Dictionary<string, Jsyd>();
            string zdnum;
            foreach (Jsyd jsyd in jsyds)
            {
                zdnum = jsyd.Zdnum;
                if (dic.ContainsKey(zdnum))
                {
                    MessageBox.Show("�ڵر����ظ���" + zdnum);
                }
                else
                {
                    dic.Add(zdnum, jsyd);
                }
            }
            return dic;
        }

        /// <summary>
        /// ����ڵ�ͼ
        /// </summary>
        /// <param name="jsyd"></param>
        /// <param name="entitys">ͼ���ڵ�����ʵ��</param>
        internal static void CheckZDT(Jsyd jsyd, Autodesk.AutoCAD.DatabaseServices.DBObjectCollection entitys)
        {
            Utils utils = new Utils();
            string layerName = "�ڵ�ͼ������";

            foreach (DBObject dBObject in entitys)
            {

                if (dBObject is DBText)
                {
                    DBText dBText2 = (DBText)dBObject;
                    string text = dBText2.TextString;
                    Point3d pt = dBText2.Position;
                    if (text.Contains("�ڵ����"))
                    {
                        if (jsyd.Zdmj != GetArea(text))
                        {
                            utils.CreateCircle(pt, 1, layerName);
                        }
                    }
                    else if (text.Contains("-"))//�ؼ�ͼ��
                    {
                        if (text != jsyd.TFH)
                        {
                            utils.CreateCircle(pt, 1, layerName);
                        }
                    }
                    else if (text.Contains("լ�������"))
                    {
                        if (jsyd.Syqmj != GetArea2(text))
                        {
                            utils.CreateCircle(pt, 1, layerName);
                        }
                    }
                    else if (text.Contains("�����õ����"))
                    {
                        if (jsyd.Czmj != GetArea2(text))
                        {
                            utils.CreateCircle(pt, 1, layerName);
                        }
                    }
                    else if (text.Contains("����Ȩ����") )
                    {
                        if (!text.Contains("��"))
                        {
                            text = text.Replace("����Ȩ���ˣ�", "").Replace("(��)", "").Replace("������", "").Trim();
                            if (jsyd.QLRMC != text)
                            {
                                utils.CreateCircle(pt, 1, layerName);
                            }
                        }
                        
                    }
                }
            }
        }
        public static double GetArea(string text)
        {
            char[] sp = new char[]
            {
                ':','��'
            };
            double num;
            string[] array = text.Split(sp);
            foreach (string str in array)
            {
                if (double.TryParse(str.Trim(), out num))
                {
                    return num;
                }
            }
            return 0;
        }
        public static double GetArea2(string text)
        {
            text = text.Replace("ʹ��Ȩ�������׼����լ�������", "").Replace("�������彨���õ����", "").Replace("ƽ����", "").Trim();
            double num;
            if (double.TryParse(text, out num))
            {
                return num;
            }
            return 0;
        }
    }
}
