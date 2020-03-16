using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace YanBo_CG
{
    internal class Utils
    {
        private static Regex numReg = new Regex("^[0-9]*$");

        private static Regex doubleReg = new Regex("(^[0-9]*[1-9][0-9]*$)|(^([0-9]{1,}[.][0-9]*)$)");

        private static Regex zjhmReg = new Regex("\\d{17}[0-9X]");

        internal static Dictionary<string, int> fontSizeDic;

        private Utils()
        {
            Utils.GetFontSizeDic();
        }

        private static Dictionary<string, int> GetFontSizeDic()
        {
            Dictionary<string, int> result;
            if (Utils.fontSizeDic != null)
            {
                result = Utils.fontSizeDic;
            }
            else
            {
                Utils.fontSizeDic = new Dictionary<string, int>();
                Utils.fontSizeDic.Add("四号", 14);
                Utils.fontSizeDic.Add("小四", 12);
                result = Utils.fontSizeDic;
            }
            return result;
        }

        public static bool IsDouble(string str)
        {
            return Utils.doubleReg.IsMatch(str);
        }

        public static bool IsNum(string str)
        {
            return !Utils.IsStrNull(str) && Utils.numReg.IsMatch(str);
        }

        public static string ListToStr(string[] array, string connect)
        {
            string text = "";
            int num = array.Length;
            for (int i = 0; i < num; i++)
            {
                if (i != 0)
                {
                    text += connect;
                }
                text += array[i];
            }
            return text;
        }

        public static string ListToStr(IList<string> list, string connect)
        {
            string text = "";
            int count = list.Count;
            for (int i = 0; i < count; i++)
            {
                if (i != 0)
                {
                    text += connect;
                }
                text += list[i];
            }
            return text;
        }

        public static StringBuilder ListToStringBuilder(IList<string> list, string connect)
        {
            StringBuilder stringBuilder = new StringBuilder();
            int count = list.Count;
            for (int i = 0; i < count; i++)
            {
                if (i != 0)
                {
                    stringBuilder.Append(connect);
                }
                stringBuilder.Append(list[i]);
            }
            return stringBuilder;
        }

        internal static void DicAdd<T>(string key, T value, Dictionary<string, IList<T>> dic)
        {
            IList<T> list;
            if (dic.TryGetValue(key, out list))
            {
                list.Add(value);
            }
            else
            {
                list = new List<T>();
                list.Add(value);
                dic.Add(key, list);
            }
        }

        internal static int GetBigNum(int num1, int num2)
        {
            return (num1 > num2) ? num1 : num2;
        }

        internal static void IListAddRange<T>(IList<T> list, IList<T> list2)
        {
            if (list2 != null)
            {
                foreach (T current in list2)
                {
                    list.Add(current);
                }
            }
        }

        public static void ListToText(List<string> list, string fileName)
        {
            FileStream fileStream = File.Create(fileName);
            StreamWriter streamWriter = new StreamWriter(fileStream);
            foreach (string current in list)
            {
                streamWriter.WriteLine(current);
            }
            streamWriter.Close();
            fileStream.Close();
        }

        internal static void RemoveList<T>(IList<T> fileinfos, IList<T> removes)
        {
            foreach (T current in removes)
            {
                fileinfos.Remove(current);
            }
        }

        internal static string StringAppend(string mfwly, string fwly, string symbol)
        {
            if (!Utils.IsStrNull(fwly))
            {
                if (Utils.IsStrNull(mfwly))
                {
                    mfwly = fwly;
                }
                else if (!mfwly.Contains(fwly))
                {
                    mfwly = mfwly + symbol + fwly;
                }
            }
            return mfwly;
        }

        public static DateTime FormatDate(string dateTime)
        {
            DateTimeFormatInfo dateTimeFormatInfo = new DateTimeFormatInfo();
            if (dateTime.Contains("年"))
            {
                dateTimeFormatInfo.ShortDatePattern = "yyyy年MM月dd日";
            }
            else if (dateTime.Contains("-"))
            {
                dateTimeFormatInfo.ShortDatePattern = "yyyy-MM-dd";
            }
            else if (dateTime.Length == 8)
            {
                dateTime = dateTime.Insert(4, "/");
                dateTime = dateTime.Insert(7, "/");
                dateTimeFormatInfo.ShortDatePattern = "yyyyMMdd";
            }
            return Convert.ToDateTime(dateTime, dateTimeFormatInfo);
        }

        public static bool IsStrNull(string str)
        {
            return str == null || str.Trim().Equals("");
        }

        public static bool StrThanLength(string str, int thanLength)
        {
            return str != null && str.Length >= thanLength;
        }

        public static bool StrEquals(string str1, string str2)
        {
            bool result;
            if (!Utils.IsStrNull(str1))
            {
                if (!str1.Equals(str2))
                {
                    result = false;
                    return result;
                }
            }
            else if (!Utils.IsStrNull(str2))
            {
                result = false;
                return result;
            }
            result = true;
            return result;
        }

        public static void RepList(IList<string> list)
        {
            if (list != null && list.Count != 0)
            {
                for (int i = 0; i < list.Count - 1; i++)
                {
                    string str = list[i];
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        string str2 = list[j];
                        if (Utils.StrEquals(str, str2))
                        {
                            list.RemoveAt(j);
                            j--;
                        }
                    }
                }
            }
        }

        internal static string FormatDate(DateTime dt)
        {
            string result;
            if (dt.Year != 1)
            {
                result = dt.ToString("yyyy年MM月dd日");
            }
            else
            {
                result = "";
            }
            return result;
        }

        public static IList<T> ArrayToList<T>(T[] array)
        {
            IList<T> list = new List<T>();
            for (int i = 0; i < array.Length; i++)
            {
                T item = array[i];
                list.Add(item);
            }
            return list;
        }

        internal static void ArrayReplaceAfter(string[] addressArray)
        {
            if (addressArray != null)
            {
                for (int i = addressArray.Length - 1; i > 0; i--)
                {
                    string text = addressArray[i];
                    string text2 = addressArray[i - 1];
                    if (text != null && text2 != null)
                    {
                        addressArray[i] = text.Replace(text2, "");
                    }
                }
            }
        }

        public static Dictionary<int, string> DicToInt(Dictionary<string, string> dic)
        {
            Dictionary<int, string> dictionary = new Dictionary<int, string>();
            Dictionary<string, string>.KeyCollection keys = dic.Keys;
            foreach (string current in keys)
            {
                if (Utils.IsNum(current))
                {
                    string value;
                    dic.TryGetValue(current, out value);
                    dictionary.Add(int.Parse(current), value);
                }
            }
            return dictionary;
        }

        internal static IList<string> GetKeys(Dictionary<string, string> dic)
        {
            IList<string> list = new List<string>();
            foreach (string current in dic.Keys)
            {
                list.Add(current);
            }
            return list;
        }

        internal static string[] ArrayRemoveNull(string[] addressArray)
        {
            int i;
            for (i = addressArray.Length - 1; i >= 0; i--)
            {
                if (addressArray[i] != null)
                {
                    break;
                }
            }
            return Utils.ArrayCopyToIndex<string>(addressArray, i + 1);
        }

        private static T[] ArrayCopyToIndex<T>(T[] array, int v)
        {
            T[] array2 = new T[v];
            for (int i = 0; i < v; i++)
            {
                array2[i] = array[i];
            }
            return array2;
        }

        public static Dictionary<string, string> keyToLower(Dictionary<string, string> dic)
        {
            Dictionary<string, string>.KeyCollection keys = dic.Keys;
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            foreach (string current in keys)
            {
                string value;
                dic.TryGetValue(current, out value);
                dictionary.Add(current.ToLower(), value);
            }
            return dictionary;
        }

        public static Dictionary<string, string> ValueToLower(Dictionary<string, string> dic)
        {
            Dictionary<string, string>.KeyCollection keys = dic.Keys;
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            foreach (string current in keys)
            {
                string text;
                dic.TryGetValue(current, out text);
                if (current.Contains("类的全名称"))
                {
                    dictionary.Add(current, text);
                }
                else
                {
                    dictionary.Add(current, text.ToLower());
                }
            }
            return dictionary;
        }

        internal static void DicAddDic<T1, T2>(Dictionary<T1, T2> map1, Dictionary<T1, T2> map2)
        {
            foreach (T1 current in map2.Keys)
            {
                if (!map1.ContainsKey(current))
                {
                    T2 value;
                    map2.TryGetValue(current, out value);
                    map1.Add(current, value);
                }
            }
        }

        internal static string GetCompany()
        {
            string result;
            if (!LoadFile.fileDic.TryGetValue("公司名称", out result))
            {
                MessageBox.Show("请在文件配置中设置‘公司名称’");
            }
            return result;
        }

        internal static string ListToText<T>(IList<T> list, string connect)
        {
            T t = list[0];
            string text = t.ToString();
            for (int i = 1; i < list.Count; i++)
            {
                text = text + connect + list[i];
            }
            return text;
        }

        internal static double AddNum(IList<double> nums)
        {
            double num = 0.0;
            foreach (double num2 in nums)
            {
                num += num2;
            }
            return num;
        }
    }
}
