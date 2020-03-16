using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows.Forms;
namespace YanBo
{
    class ExcelUtils
    {
        public String CreateJzxQZ_Table(List<Jzx> jzxList)
        {
            Char[] ch = new char[]{
               '_','-'
           };

            foreach (Jzx jzx in jzxList)
            {
                HSSFWorkbook workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\��ַǩ�±�ģ��.xls", FileMode.Open, FileAccess.Read));
                ISheet isheet = workbook[0];
                List<String> jzxSm = jzx.getJzxSm();

                for (int a = 0; a < jzxSm.Count; a++)
                {
                    String[] smArray = jzxSm[a].Split(ch);
                    if (smArray.Length == 1)
                    {
                        continue;
                    }
                    IRow irow = isheet.GetRow(a + 6);
                    if (irow == null)
                    {
                        continue;
                    }
                    //ѭ����

                    for (int b = 1; b < 8; b++)
                    {
                        ICell icell = irow.GetCell(b);
                        if (icell == null)
                        {
                            continue;
                        }
                        switch (b)
                        {
                            case 1:
                                icell.SetCellValue(smArray[0]);
                                break;
                            case 2:
                                icell.SetCellValue(smArray[1]);
                                break;
                            case 3:
                                if (smArray.Length > 2)
                                {
                                    if (smArray[2].Length > 18)
                                    {
                                        smArray[2] = smArray[2].Substring(0, smArray[2].Length - 19);
                                    }
                                    icell.SetCellValue(smArray[2]);
                                }

                                break;

                            default:
                                break;
                        }
                    }
                }
                String path = @"d:\��ַǩ�±�\";

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);

                }

                FileStream file = new FileStream(path + jzx.getCbfbm() + jzx.getCbfmc() + "_ǩ�±�.xls", FileMode.Create);
                workbook.Write(file);

                file.Close();  //�ر��ļ���  
            }

            return "";
        }

        /**
        jzxLb.Add("��ַ��", 7);
        jzxLb.Add("ǽ��", 8);
        jzxLb.Add("��ˮ", 9);
        jzxLb.Add("Χǽ", 10);
        jzxLb.Add("��·", 11);
        jzxLb.Add("����", 12);
        jzxLb.Add("��", 13);
        jzxWz.Add("��", 15);
        jzxWz.Add("��", 16);
        jzxWz.Add("��", 17);
         * **/

        public String CreateJzxBS_Table(List<Jzx> jzxList)
        {

            Dictionary<String, String> jzxBSCorresponding = ExcelUtils.ReadExcelToDic(@"C:\ZJDTemplete\��ַ��ʾ���Ӧλ��.xls", 0);

            string partent = @"d:\��ַ��ʾ��\";
            int jzxCount = 0;
            foreach (Jzx jzx in jzxList)
            {
                List<String> jzxBs = jzx.getJzxBs();

                HSSFWorkbook workbook = null;
                ISheet isheet = null;

                string outStr = "";
                jzxBSCorresponding.TryGetValue("��ʼ��", out outStr);
                String pageConut;
                jzxBSCorresponding.TryGetValue("ÿҳ��ʾ��ַ�����", out pageConut);
                int page = int.Parse(pageConut);
                int rowStart = int.Parse(outStr);
                //����
                if (jzxCount > 10)
                {
                    //ÿ�ڽ���3��
                    //System.Threading.Thread.Sleep(2000);
                }
                jzxCount++;
                for (int a = 0; a < jzxBs.Count / page + 1; a++)
                {

                    workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\��ַ������ģ��.xls", FileMode.Open, FileAccess.Read));
                    isheet = workbook[0];
                    if (jzxBs.Count < page)
                    {

                        WriteBSB_Page(isheet, rowStart, jzxBs, 0, jzxBs.Count, jzxBSCorresponding);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_��ʾ��.xls", workbook);
                    }
                    else
                    {
                        WriteBSB_Page(isheet, rowStart, jzxBs, a * (page - 1), (a + 1) * (page - 1), jzxBSCorresponding);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_" + (a + 1) + "��ʾ��.xls", workbook);

                    }

                }


            }

            return "";
        }

        public void CreateJzxBS_Table_X(List<Jzx> jzxList)
        {
            Dictionary<String, String> jzxBSCorresponding = ExcelUtils.ReadExcelToDic(@"C:\ZJDTemplete\��ַ��ʾ���Ӧλ��.xls", 0);
            string partent = @"d:\��ַ��ʾ��\";
            int jzxCount = 0;
            foreach (Jzx jzx in jzxList)
            {
                List<String> jzxBs = jzx.getJzxBs();

                HSSFWorkbook workbook = null;
                ISheet isheet = null;

                string outStr = "";
                jzxBSCorresponding.TryGetValue("��ʼ��", out outStr);
                int rowStart = int.Parse(outStr);

                String pageConut;
                jzxBSCorresponding.TryGetValue("ÿҳ��ʾ��ַ�����", out pageConut);
                int page = int.Parse(pageConut);

                //����
                if (jzxCount > 10)
                {
                    //ÿ�ڽ���3��
                    //System.Threading.Thread.Sleep(2000);
                }
                jzxCount++;
                for (int a = 0; a < jzxBs.Count / page + 1; a++)
                {

                    workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\��ַ������ģ��.xls", FileMode.Open, FileAccess.Read));
                    isheet = workbook[0];
                    if (jzxBs.Count < page)
                    {

                        WriteBSB_Page_X(isheet, rowStart, jzxBs, 0, jzxBs.Count);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_��ʾ��.xls", workbook);
                    }
                    else
                    {
                        WriteBSB_Page_X(isheet, rowStart, jzxBs, a * (page - 1), (a + 1) * (page - 1));
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_" + (a + 1) + "��ʾ��.xls", workbook);

                    }

                }


            }

        }
        public void WriterBsbExcel(String partent, String saveName, HSSFWorkbook workbook)
        {

            //����Ŀ¼
            if (!Directory.Exists(partent))
            {
                Directory.CreateDirectory(partent);

            }

            FileStream file = new FileStream(partent + saveName, FileMode.Create);
            workbook.Write(file);

            file.Close();  //�ر��ļ���  

        }
        //��ʾ����дҳ��
        public int WriteBSB_Page(ISheet isheet, int rowIndex, List<String> jzxBs, int startJzxIndex, int endJzxIndex, Dictionary<String, String> zb)
        {
            IRow irow;
            ICell icell;
            String cellIndex;

            if (endJzxIndex > jzxBs.Count)
            {
                endJzxIndex = jzxBs.Count;
            }

            for (int a = startJzxIndex; a < endJzxIndex; a++)
            {


                String bs = jzxBs[a];
                String[] bsArray = bs.Split('_');
                irow = isheet.GetRow(rowIndex++);
                icell = irow.GetCell(0);
                icell.SetCellValue(bsArray[0]);
                icell = irow.GetCell(3);
                icell.SetCellValue("�̡�");


                irow = isheet.GetRow(rowIndex++);
                icell = irow.GetCell(6);
                double num = 0;
                double.TryParse(bsArray[4], out num);
                icell.SetCellValue(num);

                zb.TryGetValue(bsArray[2], out  cellIndex);
                icell = irow.GetCell(int.Parse(cellIndex));
                icell.SetCellValue("�̡�");

                zb.TryGetValue(bsArray[3], out  cellIndex);
                icell = irow.GetCell(int.Parse(cellIndex));
                icell.SetCellValue("�̡�");
            }
            irow = isheet.GetRow(rowIndex);
            icell = irow.GetCell(0);
            icell.SetCellValue(jzxBs[endJzxIndex - 1].Split('_')[1]);
            icell = irow.GetCell(3);
            icell.SetCellValue("�̡�");



            try
            {
            }
            catch
            {
                throw new Exception("���ñ�������");
            }
            return rowIndex;
        }

        //��ʾ����дҳ��
        public void WriteBSB_Page_X(ISheet isheet, int rowIndex, List<String> jzxBs, int startJzxIndex, int endJzxIndex)
        {
            JzxService jzxService = new JzxService();
            Dictionary<String, String> dic = jzxService.GetjzxBSCorresponding();
            Dictionary<String, String> zlDic = jzxService.GetJzxZLDic();
            String zl = "";
            IRow irow;
            ICell icell;
            //int cellIndex = 0;
            String outStr = "";
            if (endJzxIndex > jzxBs.Count)
            {
                endJzxIndex = jzxBs.Count;
            }
            for (int a = startJzxIndex; a < endJzxIndex; a++)
            {
                String bs = jzxBs[a];
                String[] bsArray = bs.Split('_');
                if (bsArray[2].Equals("��ַ��"))
                {
                    bsArray[3] = "��";
                }
                irow = isheet.GetRow(rowIndex++);
                dic.TryGetValue("��ַ���", out outStr);
                icell = irow.GetCell(int.Parse(outStr));


                zlDic.TryGetValue(bsArray[2], out zl);
                zl = getJzxX_Jl(zl, jzxBs, a, zlDic);
                if (zl == null)
                {
                    MessageBox.Show("��������������⣺" + bsArray[2]);
                    continue;
                }
                dic.TryGetValue(zl, out outStr);

                icell.SetCellValue(bsArray[0]);
                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("�̡�");


                dic.TryGetValue("��ַ����", out outStr);
                irow = isheet.GetRow(rowIndex++);
                icell = irow.GetCell(int.Parse(outStr));
                double num = 0;
                double.TryParse(bsArray[4], out num);
                icell.SetCellValue(num);

                dic.TryGetValue(bsArray[2], out outStr);


                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("�̡�");

                dic.TryGetValue(bsArray[3], out outStr);

                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("�̡�");
            }
            irow = isheet.GetRow(rowIndex);
            icell = irow.GetCell(0);
            icell.SetCellValue(jzxBs[endJzxIndex - 1].Split('_')[1]);

            zlDic.TryGetValue(jzxBs[endJzxIndex - 1].Split('_')[2], out zl);
            zl = getJzxX_Jl(zl, jzxBs, endJzxIndex - 1, zlDic);
            if (zl == null)
            {

                return;
            }
            dic.TryGetValue(zl, out outStr);
            icell = irow.GetCell(int.Parse(outStr));
            icell.SetCellValue("�̡�");

        }
        public String getJzxX_Jl(String zl, List<String> jzxBs, int a, Dictionary<String, String> zlDic)
        {

            if (zl.Equals("����"))
            {
                //��Զ���ǵ�һ���㣬�鿴��һ����
                //�鿴�Ƿ��ǵ�һ����
                if (a == 0)
                {
                    //�鿴���һ�����ǲ�������
                    zlDic.TryGetValue(jzxBs[jzxBs.Count - 1].Split('_')[2], out zl);
                }
                else if (a == jzxBs.Count - 1)
                {
                    //��������һ�����Ϳ���һ��
                    zlDic.TryGetValue(jzxBs[0].Split('_')[2], out zl);
                }
                else
                {
                    zlDic.TryGetValue(jzxBs[a - 1].Split('_')[2], out zl);
                }
            }
            return zl;
        }
        //�ļ�תΪdictory
        public static Dictionary<String, String> ReadExcelToDic(String excelPath, int index)
        {
            Dictionary<String, String> dic = new Dictionary<string, string>();
            HSSFWorkbook workbook = new HSSFWorkbook(new FileStream(excelPath, FileMode.Open, FileAccess.Read));
            ISheet isheet = workbook[index];
            ICell icell;
            String key;
            String value;
            foreach (IRow irow in isheet)
            {
                icell = irow.GetCell(0);
                if (icell == null)
                {
                    continue;
                }
                icell.SetCellType(CellType.String);
                key = icell.StringCellValue;
                if (key.Equals(""))
                {
                    continue;
                }
                icell = irow.GetCell(1);
                if (icell == null)
                {
                    continue;
                }
                icell.SetCellType(CellType.String);
                value = icell.StringCellValue;
                if (value.Equals(""))
                {
                    continue;
                }
                dic.Add(key, value);

            }
            return dic;
        }

    }
}
