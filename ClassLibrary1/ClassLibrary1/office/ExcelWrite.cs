using NPOI.SS.UserModel;
using System;
using System.IO;

namespace YanBo.office
{
    public class ExcelWrite
    {
        public static void Save(IWorkbook workbook, string path)
        {
            FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate);
            workbook.Write(fileStream);
            fileStream.Close();
        }

        public static void SetValue(IRow row, int cellIndex, string value)
        {
            if (row != null)
            {
                ICell cell = row.GetCell(cellIndex);
                cell.SetCellValue(value);
            }
        }

        public static void CreateRows(ISheet sheet, int modelIndex, int cellTotal, int start, int rowsCount)
        {
            sheet.ShiftRows(start + 1, sheet.LastRowNum + 1, rowsCount);
            IRow row = sheet.GetRow(start);
            for (int i = 0; i < rowsCount; i++)
            {
                IRow row2 = sheet.CreateRow(start + 1 + i);
                for (int j = 0; j < cellTotal; j++)
                {
                    ICell cell = row2.CreateCell(j);
                    ICell cell2 = row.GetCell(j);
                    if (cell2 != null)
                    {
                        cell.CellStyle = cell2.CellStyle;
                    }
                }
            }
        }
    }
}
