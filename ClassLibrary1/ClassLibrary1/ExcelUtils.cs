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
                HSSFWorkbook workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\界址签章表模板.xls", FileMode.Open, FileAccess.Read));
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
                    //循环行

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
                String path = @"d:\界址签章表\";

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);

                }

                FileStream file = new FileStream(path + jzx.getCbfbm() + jzx.getCbfmc() + "_签章表.xls", FileMode.Create);
                workbook.Write(file);

                file.Close();  //关闭文件流  
            }

            return "";
        }

        /**
        jzxLb.Add("界址线", 7);
        jzxLb.Add("墙壁", 8);
        jzxLb.Add("滴水", 9);
        jzxLb.Add("围墙", 10);
        jzxLb.Add("道路", 11);
        jzxLb.Add("沟渠", 12);
        jzxLb.Add("坎", 13);
        jzxWz.Add("内", 15);
        jzxWz.Add("中", 16);
        jzxWz.Add("外", 17);
         * **/

        public String CreateJzxBS_Table(List<Jzx> jzxList)
        {

            Dictionary<String, String> jzxBSCorresponding = ExcelUtils.ReadExcelToDic(@"C:\ZJDTemplete\界址标示表对应位置.xls", 0);

            string partent = @"d:\界址标示表\";
            int jzxCount = 0;
            foreach (Jzx jzx in jzxList)
            {
                List<String> jzxBs = jzx.getJzxBs();

                HSSFWorkbook workbook = null;
                ISheet isheet = null;

                string outStr = "";
                jzxBSCorresponding.TryGetValue("起始行", out outStr);
                String pageConut;
                jzxBSCorresponding.TryGetValue("每页显示界址点个数", out pageConut);
                int page = int.Parse(pageConut);
                int rowStart = int.Parse(outStr);
                //降速
                if (jzxCount > 10)
                {
                    //每宗降低3秒
                    //System.Threading.Thread.Sleep(2000);
                }
                jzxCount++;
                for (int a = 0; a < jzxBs.Count / page + 1; a++)
                {

                    workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\界址线类型模板.xls", FileMode.Open, FileAccess.Read));
                    isheet = workbook[0];
                    if (jzxBs.Count < page)
                    {

                        WriteBSB_Page(isheet, rowStart, jzxBs, 0, jzxBs.Count, jzxBSCorresponding);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_标示表.xls", workbook);
                    }
                    else
                    {
                        WriteBSB_Page(isheet, rowStart, jzxBs, a * (page - 1), (a + 1) * (page - 1), jzxBSCorresponding);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_" + (a + 1) + "标示表.xls", workbook);

                    }

                }


            }

            return "";
        }

        public void CreateJzxBS_Table_X(List<Jzx> jzxList)
        {
            Dictionary<String, String> jzxBSCorresponding = ExcelUtils.ReadExcelToDic(@"C:\ZJDTemplete\界址标示表对应位置.xls", 0);
            string partent = @"d:\界址标示表\";
            int jzxCount = 0;
            foreach (Jzx jzx in jzxList)
            {
                List<String> jzxBs = jzx.getJzxBs();

                HSSFWorkbook workbook = null;
                ISheet isheet = null;

                string outStr = "";
                jzxBSCorresponding.TryGetValue("起始行", out outStr);
                int rowStart = int.Parse(outStr);

                String pageConut;
                jzxBSCorresponding.TryGetValue("每页显示界址点个数", out pageConut);
                int page = int.Parse(pageConut);

                //降速
                if (jzxCount > 10)
                {
                    //每宗降低3秒
                    //System.Threading.Thread.Sleep(2000);
                }
                jzxCount++;
                for (int a = 0; a < jzxBs.Count / page + 1; a++)
                {

                    workbook = new HSSFWorkbook(new FileStream(@"C:\ZJDTemplete\界址线类型模板.xls", FileMode.Open, FileAccess.Read));
                    isheet = workbook[0];
                    if (jzxBs.Count < page)
                    {

                        WriteBSB_Page_X(isheet, rowStart, jzxBs, 0, jzxBs.Count);
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_标示表.xls", workbook);
                    }
                    else
                    {
                        WriteBSB_Page_X(isheet, rowStart, jzxBs, a * (page - 1), (a + 1) * (page - 1));
                        WriterBsbExcel(partent, jzx.getCbfbm() + jzx.getCbfmc() + "_" + (a + 1) + "标示表.xls", workbook);

                    }

                }


            }

        }
        public void WriterBsbExcel(String partent, String saveName, HSSFWorkbook workbook)
        {

            //创建目录
            if (!Directory.Exists(partent))
            {
                Directory.CreateDirectory(partent);

            }

            FileStream file = new FileStream(partent + saveName, FileMode.Create);
            workbook.Write(file);

            file.Close();  //关闭文件流  

        }
        //标示表填写页面
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
                icell.SetCellValue("√　");


                irow = isheet.GetRow(rowIndex++);
                icell = irow.GetCell(6);
                double num = 0;
                double.TryParse(bsArray[4], out num);
                icell.SetCellValue(num);

                zb.TryGetValue(bsArray[2], out  cellIndex);
                icell = irow.GetCell(int.Parse(cellIndex));
                icell.SetCellValue("√　");

                zb.TryGetValue(bsArray[3], out  cellIndex);
                icell = irow.GetCell(int.Parse(cellIndex));
                icell.SetCellValue("√　");
            }
            irow = isheet.GetRow(rowIndex);
            icell = irow.GetCell(0);
            icell.SetCellValue(jzxBs[endJzxIndex - 1].Split('_')[1]);
            icell = irow.GetCell(3);
            icell.SetCellValue("√　");



            try
            {
            }
            catch
            {
                throw new Exception("配置表有问题");
            }
            return rowIndex;
        }

        //标示表填写页面
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
                if (bsArray[2].Equals("界址线"))
                {
                    bsArray[3] = "中";
                }
                irow = isheet.GetRow(rowIndex++);
                dic.TryGetValue("界址点号", out outStr);
                icell = irow.GetCell(int.Parse(outStr));


                zlDic.TryGetValue(bsArray[2], out zl);
                zl = getJzxX_Jl(zl, jzxBs, a, zlDic);
                if (zl == null)
                {
                    MessageBox.Show("你配的数据有问题：" + bsArray[2]);
                    continue;
                }
                dic.TryGetValue(zl, out outStr);

                icell.SetCellValue(bsArray[0]);
                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("√　");


                dic.TryGetValue("界址距离", out outStr);
                irow = isheet.GetRow(rowIndex++);
                icell = irow.GetCell(int.Parse(outStr));
                double num = 0;
                double.TryParse(bsArray[4], out num);
                icell.SetCellValue(num);

                dic.TryGetValue(bsArray[2], out outStr);


                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("√　");

                dic.TryGetValue(bsArray[3], out outStr);

                icell = irow.GetCell(int.Parse(outStr));
                icell.SetCellValue("√　");
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
            icell.SetCellValue("√　");

        }
        public String getJzxX_Jl(String zl, List<String> jzxBs, int a, Dictionary<String, String> zlDic)
        {

            if (zl.Equals("其他"))
            {
                //永远都是第一个点，查看上一条线
                //查看是否是第一个点
                if (a == 0)
                {
                    //查看最后一个线是不是其他
                    zlDic.TryGetValue(jzxBs[jzxBs.Count - 1].Split('_')[2], out zl);
                }
                else if (a == jzxBs.Count - 1)
                {
                    //如果是最后一个，就看第一个
                    zlDic.TryGetValue(jzxBs[0].Split('_')[2], out zl);
                }
                else
                {
                    zlDic.TryGetValue(jzxBs[a - 1].Split('_')[2], out zl);
                }
            }
            return zl;
        }
        //文件转为dictory
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
