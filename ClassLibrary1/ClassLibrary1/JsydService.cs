using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using YanBo_CG.ExcelUtils;

namespace YanBo
{
	public class JsydService
	{
		internal static Dictionary<string, Jsyd> ExcelToJsyd(string path)
		{
			Dictionary<string, Jsyd> dictionary = new Dictionary<string, Jsyd>();
			IWorkbook workbook = ExcelRead.ReadExcel(path);
			ISheet sheetAt = workbook.GetSheetAt(0);
			for (int i = 4; i <= sheetAt.LastRowNum; i++)
			{
				Jsyd jsyd = new Jsyd();
				IRow row = sheetAt.GetRow(i);
				if (row != null)
				{
					for (int j = 0; j < 50; j++)
					{
						ICell cell = row.GetCell(j);
						if (cell != null && cell.CellType != CellType.Blank)
						{
							cell.SetCellType(CellType.String);
							string text = cell.StringCellValue.Trim();
							int num = j;
							if (num != 1)
							{
								switch (num)
								{
								case 30:
								{
									double num2;
									if (double.TryParse(text, out num2))
									{
										jsyd.Zdmj = num2;
									}
									break;
								}
								case 31:
								{
									double num2;
									if (double.TryParse(text, out num2))
									{
										jsyd.Syqmj = num2;
									}
									break;
								}
								case 34:
								{
									double num2;
									if (double.TryParse(text, out num2))
									{
										jsyd.Czmj = num2;
									}
									break;
								}
								}
							}
							else
							{
								jsyd.Zdnum = text;
								if (text.Length == 19)
								{
									dictionary.Add(text, jsyd);
								}
							}
						}
					}
				}
			}
			return dictionary;
		}
	}
}
