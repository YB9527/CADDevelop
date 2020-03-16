using System;
using System.Collections.Generic;

using System.Text;
using System.IO;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using YanBo;


namespace YanBo_CG.ExcelUtils
{
    public class ExcelRead
    {

        private static Dictionary<string, Dictionary<string, string>> dicCache;
        private static Dictionary<String, ISheet> sheetCache;
        public static IWorkbook ReadExcel(string fileName)
        {
            if (Utils.IsStrNull(fileName))
            {
                throw new Exception("文件没有选择路径");
            }
            if (!File.Exists(fileName))
            {
                throw new Exception("文件不存在：" + fileName);
            }
            fileName = fileName.Trim();
            IWorkbook workbook = null;
     
            //FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            }
            fileStream.Close();
            return workbook;
        }


        public static ISheet ReadExcelSheet(String fileName, int sheetIndex)
        {

            IWorkbook workbook = ReadExcel(fileName);
            if (workbook.NumberOfSheets <= sheetIndex)
            {
                throw new Exception(fileName+",文件没有第"+(sheetIndex+1)+"页");
            }
            ISheet sheet = workbook.GetSheetAt(sheetIndex);
            return sheet;
        }
        /// <summary>
        /// 直接转换 文件配置中的 workbook
        /// </summary>
        /// <param name="fileNameValue"></param>
        /// <returns></returns>
        public static IWorkbook FileConfigToWorkbook(String fileNameValue)
        {
            String fileName;
            if ((LoadFile.fileDic.TryGetValue(fileNameValue, out fileName)))
            {
                return ReadExcel(fileName);
            }
            else
            {
                throw new Exception("文件配置中没有配置 : " + fileNameValue);
            }
        }
        /// <summary>
        /// 直接转换 文件配置中的 sheet
        /// </summary>
        /// <param name="fileNameValue"></param>
        /// <returns></returns>
        public static ISheet FileConfigToSheet(String fileNameValue,int sheetIndex)
        {
            String fileName;
            if ((LoadFile.fileDic.TryGetValue(fileNameValue, out fileName)))
            {
                return ReadExcelSheet(fileName, sheetIndex);
            }
            else
            {
                throw new Exception("文件配置中没有配置 : " + fileNameValue);
            }
        }
        /// <summary>
        ///配置中的文件 转换dic
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static Dictionary<String, String> FileConfigToDic(String fileName, int index)
        {
            ISheet sheet = FileConfigToSheet(fileName,0);
            return SheetToDic(sheet);
        }


        /// <summary>
        /// sheet 转换dic
        /// </summary>
        /// <param name="isheet"></param>
        /// <returns></returns>
        public static Dictionary<String, String> SheetToDic(ISheet isheet)
        {
            Dictionary<String, String> dic = new Dictionary<string, string>();
            ICell icell;
            String key;
            String value;
            String outer;
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
                if (!dic.TryGetValue(key, out outer))
                {
                    dic.Add(key, value);
                }
            }
            return dic;
        }
        public static Dictionary<String, String> SheetToDic(ISheet isheet,int startRow)
        {
            Dictionary<String, String> dic = new Dictionary<string, string>();
            ICell icell;
            String key;
            String value;
            String outer;
            for (int a = startRow; a <= isheet.LastRowNum; a++)
            {
                IRow irow = isheet.GetRow(a);
                if(irow == null)
                {
                    continue;
                }
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
                if (!dic.TryGetValue(key, out outer))
                {
                    dic.Add(key, value);
                }
            }
            return dic;
        }
        //文件转为dictory
        public static Dictionary<String, String> ReadExcelToDic(String excelPath, int index)
        {
         
            ISheet sheet = ReadExcelSheet(excelPath, index);

            return SheetToDic(sheet);
        }
        public static Dictionary<String, String> ReadExcelToDic(String excelPath, int index, int startRow)
        {

            ISheet sheet = ReadExcelSheet(excelPath, index);

            return SheetToDic(sheet, startRow);
        }
        internal ICellStyle GetCellStyle(ISheet sheet, int rowIndexStyle, int CellIndexStyle)
        {
            IRow row = sheet.GetRow(rowIndexStyle);
            ICellStyle style = row.GetCell(CellIndexStyle).CellStyle;
            return style;
        }

        /// <summary>
        /// select 1 是开始 ，2包括，3结尾，4相等
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="value"></param>
        /// <param name="select"></param>
        /// <returns></returns>
        internal int[] FindContainsText(ISheet sheet, string value, int select)
        {

            int[] positon = new int[2];
            foreach (IRow row in sheet)
            {
               
                foreach (ICell cell in row)
                {
                    if (cell.CellType != CellType.Blank)
                    {
                        cell.SetCellType(CellType.String);

                        switch (select)
                        {
                            case 1:
                                if (cell.StringCellValue.Trim().StartsWith(value))
                                {
                                    positon[0] = row.RowNum;
                                    positon[1] = cell.ColumnIndex;
                                    return positon;
                                }
                                break;

                            case 2:
                                if (cell.StringCellValue.Trim().Contains(value))
                                {
                                    positon[0] = row.RowNum;
                                    positon[1] = cell.ColumnIndex;
                                    return positon;
                                }
                                break;

                            case 3:
                                if (cell.StringCellValue.Trim().EndsWith(value))
                                {
                                    positon[0] = row.RowNum;
                                    positon[1] = cell.ColumnIndex;
                                    return positon;
                                }
                                break;
                            case 4:
                                if (cell.StringCellValue.Trim().Equals(value))
                                {
                                    positon[0] = row.RowNum;
                                    positon[1] = cell.ColumnIndex;
                                    return positon;
                                }
                                break;


                        }
                    }
                }
            }
            return null;
        }




        private static Dictionary<string, Dictionary<string, string>> GetCacheDic()
        {
            if (dicCache == null)
            {
                dicCache = new Dictionary<string, Dictionary<string, string>>();
            }
            return dicCache;
        }

        //文件转为dictory
        public static Dictionary<int, String> ReadExcelToDic(String excelPath)
        {
            Dictionary<int, String> dic = new Dictionary<int, String>();

            ISheet isheet = ReadExcelSheet(excelPath, 0);
            ICell icell;
            int key;
            String value;
            String outer;
            foreach (IRow irow in isheet)
            {
                icell = irow.GetCell(0);
                if (icell == null)
                {
                    continue;
                }
                icell.SetCellType(CellType.String);
                try
                {
                    key = int.Parse(icell.StringCellValue);
                }
                catch
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
                if (!dic.TryGetValue(key, out outer))
                {
                    if (dic.ContainsKey(key))
                    {
                        throw new Exception("表中值重复：" + key);
                    }
                    dic.Add(key, value);
                }


            }
            return dic;
        }


        internal static ISheet ReadExcelSheet(String fileName, string sheetName)
        {
            IWorkbook workbook = ReadExcel(fileName);
            ISheet sheet = workbook.GetSheet(sheetName);
            return sheet;
        }
        
        internal IList<string> ReadCellsStr(string fileName)
        {
            return ReadCellsStr(fileName, 0, -1);
        }
        public IList<String> ReadCellsStr(String fileName, int startRow, int endRow)
        {
            IList<String> list = new List<String>();
            ISheet sheet = ReadExcelSheet(fileName, 0);
            ICell cell;
            IRow row;
            String value;
            if (endRow == -1)
            {
                endRow = sheet.LastRowNum + 1;
            }
            for (int a = startRow; a < endRow; a++)
            {
                row = sheet.GetRow(a);
                if (row == null)
                {
                    continue;
                }
                for (int b = 0; b < row.LastCellNum + 1; b++)
                {
                    cell = row.GetCell(b);
                    if (cell == null)
                    {
                        continue;
                    }
                    cell.SetCellType(CellType.String);
                    value = cell.StringCellValue;
                    if (!Utils.IsStrNull(value))
                    {
                        list.Add(value);
                    }
                }

            }
            return list;
        }
        /// <summary>
        /// 根据configname 读取第一列的数据集合
        /// </summary>
        /// <param name="configName"></param>
        /// <returns></returns>
        internal static IList<string> FileConfigToList(string configName)
        {
            String path = LoadFile.GetPath(configName);
            return ReadExcelToList(path);
        }

        private static IList<string> ReadExcelToList(string path)
        {
            ISheet sheet = ReadExcelSheet(path, 0);
            return ReadExcelToList(sheet);
        }
        internal static IList<string> FileToList(string path)
        {
            return ReadExcelToList(ReadExcelSheet(path,0));
        }
        internal static IList<string> FileToList(string path,int cellIndex)
        {
            return ReadExcelToList(ReadExcelSheet(path, 0), cellIndex);
        }
        private static IList<string> ReadExcelToList(ISheet sheet)
        {
            IList<string> list = new List<String>();
           foreach (IRow row in sheet)
            {
                foreach(ICell cell in row)
                {
                    if(cell.CellType != CellType.Blank)
                    {
                        cell.SetCellType(CellType.String);
                        list.Add(cell.StringCellValue);
                    }
                }
            }
            return list;
        }
        private static IList<string> ReadExcelToList(ISheet sheet,int Index)
        {
            IList<string> list = new List<String>();
            foreach (IRow row in sheet)
            {
                ICell cell = row.GetCell(Index);
                if (cell != null && cell.CellType != CellType.Blank)
                {
                    cell.SetCellType(CellType.String);
                    list.Add(cell.StringCellValue);
                }else
                {
                    list.Add(null);
                }
            }
            return list;
        }
      
        /// <summary>
        /// Excek表头检查
        /// </summary>
        /// <param name="row">要检查的行号</param>
        /// <param name="fileName"></param>
        /// <param name="sheetIndex"></param>
        /// <param name="sheetIndex"></param>
        internal static StringBuilder CheckExcelHead(IRow row, string fileName, int sheetIndex, int rowIndex)
        {

            ISheet compareSheet = ReadExcelSheet(fileName, sheetIndex);
            IRow rowCompare = compareSheet.GetRow(rowIndex);
            return CheckExcelHeadRow(row, rowCompare);


        }   


        

        private static StringBuilder CheckExcelHeadRow(IRow row, IRow rowCompare)

        {
             StringBuilder sb = new StringBuilder();
            if (row == null)
            {
                sb.Append("农房表格行数据为空，请检查");
                return sb;
            }
            ICell cell;
            ICell cellCompare;
            for (int a = 0; a < row.LastCellNum + 1; a++)
            {
                cell = row.GetCell(a);
                if (cell == null)
                {
                    continue;
                }
                cell.SetCellType(CellType.String);
                String cellValue = cell.StringCellValue;
                cellCompare = rowCompare.GetCell(a);
                if (cellCompare == null)
                {
                    sb.Append("第" + (a + 1) + "列，没有数据");
                    continue;
                }
                cellCompare.SetCellType(CellType.String);
                String cellCompareValue = cellCompare.StringCellValue;
                if (!cellCompareValue.Equals(cellValue))
                {
                    sb.Append("第" + (a + 1) + "列不是：“" + cellValue
                           + "”，数据");
                    sb.Append(",是：“" + cellCompareValue
                            + "”，数据\r\n");
                }

            }
            if (sb.Length > 1)
            {
                sb.Insert(0, "表头有问题：  表的第" + (row.RowNum + 1) + "行\r\n");
            }
            return sb;
        }

        internal static StringBuilder CheckExcelHead(string filePath, int sheetIndex,int rowIndex, string configFileName, int configSheetIndex ,int configRowIndex)
        {
            IRow row = ReadExcelSheet(filePath, sheetIndex).GetRow(rowIndex);
            ISheet compareSheet = ReadExcelSheetCache(configFileName, configSheetIndex);
            return CheckExcelHeadRow(row, compareSheet.GetRow(configRowIndex));
        }

        private static ISheet ReadExcelSheetCache(string conFileName, int sheetIndex)
        {
            if(sheetCache == null)
            {
                sheetCache = new Dictionary<string, ISheet>();
            }
            ISheet sheet;
            if(!sheetCache.TryGetValue(conFileName+ sheetIndex, out sheet))
            {
                sheet = FileConfigToSheet(conFileName, sheetIndex);
                sheetCache.Add(conFileName + sheetIndex, sheet);
            }
            
            return sheet;
        }
    }
}
