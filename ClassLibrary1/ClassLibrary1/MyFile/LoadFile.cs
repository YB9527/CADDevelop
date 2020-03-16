using System;
using System.Collections.Generic;
using YanBo_CG.ExcelUtils;

namespace YanBo_CG
{
	internal class LoadFile
	{
		public static Dictionary<string, string> fileDic;

		public static Dictionary<string, string> GetfileDic(string fileName)
		{
			LoadFile.fileDic = ExcelRead.ReadExcelToDic(fileName, 0);
			return LoadFile.fileDic;
		}

		public static string GetPath(string configName)
		{
			string result;
			if (LoadFile.fileDic.TryGetValue(configName, out result))
			{
				return result;
			}
			throw new Exception("你的文件配置中没有：" + configName);
		}
	}
}
