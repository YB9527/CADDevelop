using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Reflection;
using YanBo_CG.ExcelUtils;
using YanBo;

namespace YanBo_CG.ReFlect
{
	public class ReExcelUtils
	{
		private static Dictionary<string, Dictionary<string, Clazz>> ReflectDicCache;

		private static ExcelToObject ExcelToObject = new ExcelToObject();

		public ReExcelUtils()
		{
			if (ReExcelUtils.ReflectDicCache == null)
			{
				ReExcelUtils.ReflectDicCache = new Dictionary<string, Dictionary<string, Clazz>>();
			}
		}

		public Dictionary<int, Clazz> GetExcelTitleToClazzMap(ISheet sheet, string mapFile, int titleIndex, string classForName)
		{
			Dictionary<string, int> titleList = this.getTitleList(sheet.GetRow(titleIndex));
			Dictionary<int, Clazz> result;
			if (titleList == null)
			{
				result = null;
			}
			else
			{
				Dictionary<string, string> map = ExcelRead.ReadExcelToDic(mapFile, 0);
				Dictionary<int, string> functionMap = this.MapToMapKeyToValue(titleList, map);
				Dictionary<string, Clazz> methodMap = ReExcelUtils.MethodToLowerFunction(classForName);
				Dictionary<int, Clazz> dictionary = this.MapToMapKeyToValue(functionMap, methodMap);
				result = dictionary;
			}
			return result;
		}

		public Dictionary<int, Clazz> GetExcelTitleToClazzMap(string path, string mapFile, int titleIndex, string classForName)
		{
			IWorkbook workbook = ExcelRead.ReadExcel(path);
			ISheet sheetAt = workbook.GetSheetAt(0);
			return this.GetExcelTitleToClazzMap(sheetAt, mapFile, titleIndex, classForName);
		}

		internal Dictionary<int, Clazz> ConfigNameToClazzMap(string configNamePath, int sheetIndex, object obj)
		{
			Dictionary<string, string> dic = ExcelRead.FileConfigToDic(configNamePath, sheetIndex);
			return this.DicToClazzMap(dic, obj);
		}

		internal Dictionary<int, Clazz> ConfigNameToClazzMap(string configNamePath, int sheetIndex, string classFullName)
		{
			Dictionary<string, string> dic = ExcelRead.FileConfigToDic(configNamePath, sheetIndex);
            Dictionary<int, string> functionMap = Utils.DicToInt(dic);
			return this.GetExcelToClazzMap(functionMap, classFullName);
		}

		internal Dictionary<int, Clazz> DicToClazzMap(Dictionary<string, string> dic, object obj)
		{
			Dictionary<int, string> functionMap = Utils.DicToInt(dic);
			string fullName = obj.GetType().FullName;
			return this.GetExcelToClazzMap(functionMap, fullName);
		}

		public Dictionary<int, Clazz> GetExcelToClazzMap(Dictionary<int, string> functionMap, string classForName)
		{
			functionMap.Add(-1, "objectid");
			Dictionary<string, Clazz> methodMap = ReExcelUtils.MethodToLowerFunction(classForName);
			return this.MapToMapKeyToValue(functionMap, methodMap);
		}

		private Dictionary<int, Clazz> MapToMapKeyToValue(Dictionary<int, string> functionMap, Dictionary<string, Clazz> methodMap)
		{
			Dictionary<int, Clazz> dictionary = new Dictionary<int, Clazz>();
			Dictionary<int, string>.KeyCollection keys = functionMap.Keys;
			foreach (int current in keys)
			{
				string text;
				functionMap.TryGetValue(current, out text);
				Clazz value;
				if (methodMap.TryGetValue(text.ToLower(), out value))
				{
					dictionary.Add(current, value);
				}
			}
			return dictionary;
		}

		private Dictionary<string, int> KeyToValue(Dictionary<int, string> dic)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			Dictionary<int, string>.KeyCollection keys = dic.Keys;
			foreach (int current in keys)
			{
				string key;
				if (dic.TryGetValue(current, out key))
				{
					dictionary.Add(key, current);
				}
			}
			return dictionary;
		}

		private Dictionary<int, string> MapToMapKeyToValue(Dictionary<string, int> title, Dictionary<string, string> map)
		{
			Dictionary<int, string> dictionary = new Dictionary<int, string>();
			Dictionary<string, int>.KeyCollection keys = title.Keys;
			foreach (string current in keys)
			{
				string value;
				if (map.TryGetValue(current, out value))
				{
					int key;
					title.TryGetValue(current, out key);
					dictionary.Add(key, value);
				}
			}
			return dictionary;
		}

		private Dictionary<string, Clazz> MethodToFunction(object obj)
		{
			return ReExcelUtils.MethodToLowerFunction(obj.GetType().FullName);
		}

		public static Dictionary<string, Clazz> MethodToLowerFunction(string classForName)
		{
			Dictionary<string, Clazz> dictionary = new Dictionary<string, Clazz>();
			Type type = Type.GetType(classForName);
			object obj = type.Assembly.CreateInstance(classForName);
			MethodInfo[] methods = type.GetMethods();
			MethodInfo[] array = methods;
			for (int i = 0; i < array.Length; i++)
			{
				MethodInfo methodInfo = array[i];
				ParameterInfo[] parameters = methodInfo.GetParameters();
				if (parameters.Length != 0)
				{
					string text = methodInfo.Name.Replace("set_", "");
					text = text.ToLower();
					Type parameterType = parameters[0].ParameterType;
					Clazz clazz = new Clazz();
					clazz.setFunction(text);
					clazz.setParamterType(parameterType);
					clazz.SetMethodInfo = methodInfo;
					clazz.setFullClassName(classForName);
					clazz.SetMethodInfo = methodInfo;
					dictionary.Add(text, clazz);
				}
			}
			array = methods;
			for (int i = 0; i < array.Length; i++)
			{
				MethodInfo methodInfo = array[i];
				ParameterInfo[] parameters = methodInfo.GetParameters();
				if (parameters.Length == 0)
				{
					string text = methodInfo.Name.Replace("get_", "");
					text = text.ToLower();
					Clazz clazz;
					if (dictionary.TryGetValue(text, out clazz))
					{
						clazz.GetMethodInfo = methodInfo;
					}
				}
			}
			return dictionary;
		}

		public static Dictionary<string, Clazz> MethodToFunction(string classForName)
		{
			Dictionary<string, Clazz> dictionary;
			Dictionary<string, Clazz> result;
			if (ReExcelUtils.ReflectDicCache.TryGetValue(classForName, out dictionary))
			{
				result = dictionary;
			}
			else
			{
				dictionary = new Dictionary<string, Clazz>();
				Type type = Type.GetType(classForName);
				object obj = type.Assembly.CreateInstance(classForName);
				MethodInfo[] methods = type.GetMethods();
				MethodInfo[] array = methods;
				for (int i = 0; i < array.Length; i++)
				{
					MethodInfo methodInfo = array[i];
					ParameterInfo[] parameters = methodInfo.GetParameters();
					if (parameters.Length != 0)
					{
						string text = methodInfo.Name.Replace("set_", "");
						Type parameterType = parameters[0].ParameterType;
						Clazz clazz = new Clazz();
						clazz.setFunction(text);
						clazz.setParamterType(parameterType);
						clazz.SetMethodInfo = methodInfo;
						clazz.setFullClassName(classForName);
						dictionary.Add(text, clazz);
					}
				}
				array = methods;
				for (int i = 0; i < array.Length; i++)
				{
					MethodInfo methodInfo = array[i];
					ParameterInfo[] parameters = methodInfo.GetParameters();
					if (parameters.Length == 0)
					{
						string text = methodInfo.Name.Replace("get_", "");
						Clazz clazz;
						if (dictionary.TryGetValue(text, out clazz))
						{
							clazz.GetMethodInfo = methodInfo;
						}
					}
				}
				ReExcelUtils.ReflectDicCache.Add(classForName, dictionary);
				result = dictionary;
			}
			return result;
		}

		public Dictionary<string, int> getTitleList(IRow row)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			Dictionary<string, int> result;
			if (row == null)
			{
				result = null;
			}
			else
			{
				foreach (ICell cell in row)
				{
					cell.SetCellType(CellType.String);
					string stringCellValue = cell.StringCellValue;
					if (stringCellValue != null && !stringCellValue.Equals(""))
					{
						dictionary.Add(stringCellValue, cell.ColumnIndex);
					}
				}
				result = dictionary;
			}
			return result;
		}

		public string GetFullName(Dictionary<string, Clazz> clazzDic)
		{
			Dictionary<string, Clazz>.KeyCollection keys = clazzDic.Keys;
			string result;
			using (Dictionary<string, Clazz>.KeyCollection.Enumerator enumerator = keys.GetEnumerator())
			{
				if (enumerator.MoveNext())
				{
					string current = enumerator.Current;
					Clazz clazz;
					clazzDic.TryGetValue(current, out clazz);
					string fullClassName = clazz.getFullClassName();
					result = fullClassName;
					return result;
				}
			}
			result = null;
			return result;
		}

		internal static Dictionary<string, Clazz> MethodToFunction<T>()
		{
			return ReExcelUtils.MethodToFunction(typeof(T).ToString());
		}

		public static void ClassCopy(object srcObj, object descObj)
		{
			if (srcObj != null)
			{
				Dictionary<string, Clazz> dictionary = ReExcelUtils.MethodToFunction(srcObj.GetType().ToString());
				Dictionary<string, Clazz> dictionary2 = ReExcelUtils.MethodToFunction(descObj.GetType().ToString());
				object[] array = new object[1];
				foreach (string current in dictionary.Keys)
				{
					Clazz clazz;
					if (dictionary2.TryGetValue(current, out clazz))
					{
						MethodInfo getMethodInfo = dictionary[current].GetMethodInfo;
						if (getMethodInfo != null)
						{
							array[0] = getMethodInfo.Invoke(srcObj, null);
							clazz.SetMethodInfo.Invoke(descObj, array);
						}
					}
				}
			}
		}
	}
}
