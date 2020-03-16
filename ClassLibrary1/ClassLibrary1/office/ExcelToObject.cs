using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using YanBo_CG.ReFlect;

namespace YanBo_CG.ExcelUtils
{
	internal class ExcelToObject
	{
		private static ReExcelUtils re = new ReExcelUtils();

		public Dictionary<T1, T2> ChangeValue<T1, T2>(Dictionary<T1, T1> dic, Dictionary<T1, T2> calzzDic)
		{
			Dictionary<T1, T2> dictionary = new Dictionary<T1, T2>();
			Dictionary<T1, T1>.KeyCollection keys = dic.Keys;
			foreach (T1 current in keys)
			{
				T1 key;
				dic.TryGetValue(current, out key);
				T2 value;
				if (calzzDic.TryGetValue(key, out value))
				{
					dictionary.Add(current, value);
				}
			}
			return dictionary;
		}

		public Dictionary<T1, T2> ChangeKey<T1, T2>(Dictionary<T1, T2> calzzDic, Dictionary<T1, T1> dic)
		{
			Dictionary<T1, T2> dictionary = new Dictionary<T1, T2>();
			Dictionary<T1, T1>.KeyCollection keys = dic.Keys;
			foreach (T1 current in keys)
			{
				T2 value;
				if (calzzDic.TryGetValue(current, out value))
				{
					T1 key;
					dic.TryGetValue(current, out key);
					dictionary.Add(key, value);
				}
			}
			return dictionary;
		}

		public Dictionary<int, string> valuesToLower(Dictionary<int, string> functionMap)
		{
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			Dictionary<int, string>.KeyCollection keys = functionMap.Keys;
			foreach (int current in keys)
			{
				string text;
				functionMap.TryGetValue(current, out text);
				dictionary.Add(current, text.ToLower());
			}
			return dictionary;
		}

		public object GetCellTrueType(ICell cell, string reultType)
		{
			object obj = null;
			object result;
			if (!cell.CellType.ToString().Equals(reultType))
			{
				if (reultType != null)
				{
					if (!(reultType == "String"))
					{
						if (!(reultType == "DateTime"))
						{
							if (!(reultType == "Int32"))
							{
								if (reultType == "Double")
								{
									cell.SetCellType(CellType.String);
									string stringCellValue = cell.StringCellValue;
									if (!stringCellValue.Equals(""))
									{
										if (!Regex.Match(stringCellValue, "^(-?\\d+)(\\.\\d+)?$").Success)
										{
											throw new Exception(string.Concat(new object[]
											{
												"数据第：",
												cell.RowIndex + 1,
												" 行，第：",
												cell.ColumnIndex + 1,
												" 列,数据格式不正确,我想要小数，你给我的是：",
												stringCellValue
											}));
										}
										obj = double.Parse(stringCellValue);
									}
								}
							}
							else
							{
								cell.SetCellType(CellType.String);
								string stringCellValue = cell.StringCellValue;
								if (!stringCellValue.Equals(""))
								{
									if (!Regex.Match(stringCellValue, "^-?\\d+$").Success)
									{
										throw new Exception(string.Concat(new object[]
										{
											"数据第：",
											cell.RowIndex + 1,
											" 行，第：",
											cell.ColumnIndex + 1,
											" 列,数据格式不正确,我想要整数，你给我的是：",
											stringCellValue
										}));
									}
									obj = int.Parse(stringCellValue);
								}
							}
						}
						else
						{
							try
							{
								if (cell.CellType == CellType.Numeric)
								{
									obj = cell.DateCellValue;
									result = obj;
									return result;
								}
							}
							catch
							{
							}
							cell.SetCellType(CellType.String);
							string stringCellValue = cell.StringCellValue;
							if (!stringCellValue.Equals(""))
							{
								try
								{
									DateTime dateTime = Utils.FormatDate(stringCellValue);
									obj = dateTime;
								}
								catch
								{
									throw new Exception(string.Concat(new object[]
									{
										cell.Sheet.SheetName,
										"数据第：",
										cell.RowIndex + 1,
										" 行，第：",
										cell.ColumnIndex + 1,
										" 列,数据格式不正确,我想要日期，你给我的是：",
										stringCellValue
									}));
								}
							}
						}
					}
					else
					{
						try
						{
							cell.SetCellType(CellType.String);
							string stringCellValue = cell.StringCellValue;
							if (!stringCellValue.Equals(""))
							{
								obj = stringCellValue;
							}
						}
						catch (Exception ex)
						{
							throw new Exception("请把表格的公示去掉！！！/r/n" + ex.Message);
						}
					}
				}
			}
			else
			{
				string text = cell.CellType.ToString();
				if (text != null)
				{
					if (!(text == "String"))
					{
						if (!(text == "DateTime"))
						{
							if (!(text == "Int32"))
							{
								if (text == "Double")
								{
									obj = cell.NumericCellValue;
								}
							}
							else
							{
								obj = (int)cell.NumericCellValue;
							}
						}
						else
						{
							obj = cell.DateCellValue;
						}
					}
					else
					{
						obj = cell.StringCellValue;
					}
				}
			}
			result = obj;
			return result;
		}
	}
}
