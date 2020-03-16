using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.IO;
using YanBo_CG;

namespace YanBo
{
    internal class JzxService
    {
        private char[] jzxCh = new char[]
		{
			'-',
			'_'
		};

        private char[] hCh = new char[]
		{
			'(',
			'（'
		};

        private Utils utils = new Utils();

        private ExcelUtils exUtils = new ExcelUtils();

        private JzdService jzdSrevice = new JzdService();

        private static Dictionary<string, string> jzxBSCorresponding;

        private static Dictionary<string, string> jzxZLDic;

        public List<Jzx> getJzxQzb(object[] fromText)
        {
            List<Jzx> list;
            try
            {
                list = this.createJzxList(fromText);
            }
            finally
            {
                this.utils.deleteLayerSur("删除");
            }
            List<string> jzxSm = list[0].getJzxSm();
            this.UpdateJzxSm(list);
            List<string> list2 = this.CheckJzx(list);
            if (list2.Count > 0)
            {
                FileStream fileStream = File.Create("D:/界址签章表缺数据.txt");
                StreamWriter streamWriter = new StreamWriter(fileStream);
                foreach (string current in list2)
                {
                    streamWriter.WriteLine(current);
                }
                streamWriter.Close();
                fileStream.Close();
                Application.ShowAlertDialog("在d盘：‘界址签章表缺数据.txt’");
            }
            this.Integration(list);
            this.UpdateJzxSm(list);
            this.Integration2(list);
            return list;
        }

        public void Integration2(List<Jzx> jzxList)
        {
            for (int i = 0; i < jzxList.Count; i++)
            {
                Jzx jzx = jzxList[i];
                this.MergeJzx2(jzx);
            }
        }

        public void MergeJzx2(Jzx jzx)
        {
            List<string> list = new List<string>();
            string[] array = new string[3];
            List<string> jzxSm = jzx.getJzxSm();
            string[] array2 = jzxSm[0].Split(this.jzxCh);
            for (int i = 0; i < jzxSm.Count - 1; i++)
            {
                string[] array3 = jzxSm[i + 1].Split(this.jzxCh);
                if (array3.Length == 3)
                {
                    if (!array2[2].Equals(array3[2]))
                    {
                        list.Add(string.Concat(new string[]
						{
							array2[0],
							"_",
							array3[0],
							"_",
							array2[2]
						}));
                        array2[0] = array3[0];
                        array2[2] = array3[2];
                        if (i == jzxSm.Count - 2)
                        {
                            list.Add(string.Concat(new string[]
							{
								array3[0],
								"_",
								array3[1],
								"_",
								array3[2]
							}));
                        }
                    }
                    else if (i == jzxSm.Count - 2 && list.Count != 0)
                    {
                        list.Add(string.Concat(new string[]
						{
							array2[0],
							"_",
							array3[1],
							"_",
							array2[2]
						}));
                    }
                    if (i == jzxSm.Count - 2 && list.Count == 0)
                    {
                        list.Add(string.Concat(new string[]
						{
							array2[0],
							"_",
							array3[0],
							"_",
							array2[2]
						}));
                        list.Add(string.Concat(new string[]
						{
							array3[0],
							"_",
							array3[1],
							"_",
							array3[2]
						}));
                    }
                }
            }
            jzx.setJzxSm(list);
        }

        public List<Jzx> createJzxList(object[] formText)
        {
            List<Jzx> list = new List<Jzx>();
            Editor editor = this.utils.GetEditor();
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(8, "JZD"),
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding=("\n请选择要生成四至的JZD图层权属线");
            PromptSelectionResult selection = editor.GetSelection(promptSelectionOptions, selectionFilter);
            List<Jzx> result;
            if (selection.Status == PromptStatus.Cancel)
            {
                result = list;
            }
            else
            {
                DBObjectCollection entityCollection = this.utils.GetEntityCollection(selection);
                if (entityCollection.Count == 0)
                {
                    editor.WriteMessage("\n你没有选择对象");
                    result = list;
                }
                else
                {
                    Database workingDatabase = HostApplicationServices.WorkingDatabase;
                    using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                    {
                        Dictionary<string, string> addressDic = ExcelUtils.ReadExcelToDic((string)formText[0], 0);
                        foreach (DBObject dBObject in entityCollection)
                        {
                            if (dBObject is Polyline)
                            {
                                Polyline polyline = (Polyline)transaction.GetObject(dBObject.ObjectId, OpenMode.ForWrite, true);
                                string surAddress = this.jzdSrevice.GetSurAddress(polyline, addressDic);
                                Jzx jzx = new Jzx(surAddress);
                                ResultBuffer resultBuffer = polyline.XData;
                                if (!(resultBuffer == null))
                                {
                                    resultBuffer = polyline.GetXDataForApplication("SOUTH");
                                    TypedValue[] array2 = resultBuffer.AsArray();
                                    if (array2 != null && array2.Length >= 5)
                                    {
                                        jzx.setCbfbm(array2[2].Value.ToString());
                                        jzx.setCbfmc(array2[3].Value.ToString());
                                        this.jzdSrevice.GetSurAddress(polyline, addressDic);
                                        Point3dCollection point3dCollection = new Point3dCollection();
                                        polyline.GetStretchPoints(point3dCollection);
                                        point3dCollection = this.utils.RePoint(point3dCollection);
                                        double[] dou = this.utils.FindFourPoints(point3dCollection);
                                        Point3dCollection point3dCollection2 = this.utils.createFourPoint(dou);
                                        Polyline3d polyline3d = this.utils.CreatePolye(point3dCollection2, false);
                                        int num = this.utils.PolyDiPoly(point3dCollection, point3dCollection2);
                                        int[] array3 = new int[4];
                                        if (num > 3)
                                        {
                                            int[] array4 = this.utils.FindPointsDic(point3dCollection);
                                            int[] array5 = this.utils.FindPointsDic(point3dCollection, point3dCollection2);
                                            int[][] max = new int[][]
											{
												array4,
												array5
											};
                                            int num2 = this.utils.SelectPc1PC2(max, point3dCollection);
                                            if (num2 == 0)
                                            {
                                                array3 = array4;
                                            }
                                            else if (num2 == 1)
                                            {
                                                array3 = array5;
                                            }
                                        }
                                        else
                                        {
                                            array3 = this.utils.FindPointsDic(point3dCollection, point3dCollection2);
                                        }
                                        List<string> list2 = new List<string>();
                                        list2.AddRange(this.GetJzxSm(point3dCollection, array3[0], array3[1], surAddress, polyline, (double)formText[2]));
                                        if (list2.Count != 0 && !list2[list2.Count - 1].Equals("partition"))
                                        {
                                            list2.Add("partition");
                                        }
                                        list2.AddRange(this.GetJzxSm(point3dCollection, array3[1], array3[2], surAddress, polyline, (double)formText[2]));
                                        if (list2.Count != 0 && !list2[list2.Count - 1].Equals("partition"))
                                        {
                                            list2.Add("partition");
                                        }
                                        if (array3[2] < array3[3])
                                        {
                                            list2.AddRange(this.GetJzxSm(point3dCollection, array3[2], array3[3], surAddress, polyline, (double)formText[2]));
                                            if (list2.Count != 0 && !list2[list2.Count - 1].Equals("partition"))
                                            {
                                                list2.Add("partition");
                                            }
                                        }
                                        if (array3[3] != 0)
                                        {
                                            if (array3[3] < array3[2])
                                            {
                                                list2.AddRange(this.GetJzxSm(point3dCollection, array3[2], point3dCollection.Count - 1, surAddress, polyline, (double)formText[2]));
                                                list2.AddRange(this.GetJzxSm(point3dCollection, point3dCollection.Count - 1, array3[3], surAddress, polyline, (double)formText[2]));
                                            }
                                            else
                                            {
                                                list2.AddRange(this.GetJzxSm(point3dCollection, array3[3], point3dCollection.Count - 1, surAddress, polyline, (double)formText[2]));
                                            }
                                        }
                                        else
                                        {
                                            list2.AddRange(this.GetJzxSm(point3dCollection, array3[2], point3dCollection.Count - 1, surAddress, polyline, (double)formText[2]));
                                        }
                                        list2.AddRange(this.GetSingleJzxSm(point3dCollection, point3dCollection.Count - 1, 0, surAddress, polyline, (double)formText[2]));
                                        if (list2.Count != 0 && !list2[list2.Count - 1].Equals("partition"))
                                        {
                                            list2.Add("partition");
                                        }
                                        if (this.findJzDh(list2, "J1_") == -1 || this.findJzDh(list2, "J1-") == -1)
                                        {
                                            list2.AddRange(this.GetJzxSm(point3dCollection, 0, array3[0], surAddress, polyline, (double)formText[2]));
                                            if (list2.Count != 0 && !list2[list2.Count - 1].Equals("partition"))
                                            {
                                                list2.Add("partition");
                                            }
                                        }
                                        jzx.setJzxSm(list2);
                                        list.Add(jzx);
                                    }
                                }
                            }
                        }
                        transaction.Commit();
                    }
                    result = list;
                }
            }
            return result;
        }

        public List<string> GetJzxSm(Point3dCollection points, int startLocation, int endLocation, string address, Polyline poly, double dii)
        {
            string text = "";
            List<string> list = new List<string>();
            for (int i = startLocation; i < endLocation; i++)
            {
                double num = this.utils.CalculateTwoPt(points[i], points[i + 1]);
                if (dii > 5.0)
                {
                    if (num < 3.0)
                    {
                        num = 5.0;
                    }
                    else
                    {
                        num = dii;
                    }
                }
                else
                {
                    num = dii;
                }
                for (double num2 = 0.02; num2 < num; num2 += 1.0)
                {
                    TypedValue[] array = null;
                    object[] array2 = this.utils.selectJzxSur(points[i], points[i + 1], num2, poly);
                    if (array2 != null)
                    {
                        array = (TypedValue[])array2[0];
                    }
                    if (array != null && array.Length > 3)
                    {
                        text = array[3].Value.ToString() + array[2].Value.ToString();
                    }
                    if (text.Contains("J"))
                    {
                        if (num2 > 0.2 && num2 < 5.0)
                        {
                            text = "巷道";
                        }
                        else if (num2 > 5.0)
                        {
                            text = "空地";
                        }
                    }
                    string text2 = "";
                    if (!text2.Equals(""))
                    {
                        text = text2;
                    }
                    string text3 = string.Concat(new object[]
					{
						"J",
						i + 1,
						"-J",
						i + 2,
						"_"
					});
                    if (text.Contains("JB") || text.Contains("JC"))
                    {
                        text = text3 + text;
                        break;
                    }
                    if (dii < 0.05)
                    {
                        text = text3 + address;
                        break;
                    }
                    if (!text.Equals(""))
                    {
                        text = text3 + address + text;
                        break;
                    }
                    if (num2 >= num - 1.0)
                    {
                        text = text3;
                    }
                }
                list.Add(text);
                text = "";
            }
            return list;
        }

        public List<string> GetSingleJzxSm(Point3dCollection points, int startLocation, int endLocation, string address, Polyline poly, double dii)
        {
            string text = "";
            List<string> list = new List<string>();
            double num = this.utils.CalculateTwoPt(points[startLocation], points[endLocation]);
            if (dii > 5.0)
            {
                if (num < 3.0)
                {
                    num = 5.0;
                }
                else
                {
                    num = dii;
                }
            }
            else
            {
                num = dii;
            }
            string text2 = string.Concat(new object[]
			{
				"J",
				startLocation + 1,
				"-J",
				endLocation + 1,
				"_"
			});
            for (double num2 = 0.02; num2 < num; num2 += 1.0)
            {
                TypedValue[] array = null;
                object[] array2 = this.utils.selectJzxSur(points[startLocation], points[endLocation], num2, poly);
                if (array2 != null)
                {
                    array = (TypedValue[])array2[0];
                }
                string text3 = "";
                if (array != null && array.Length > 3)
                {
                    text3 = array[3].Value.ToString() + array[2].Value.ToString();
                }
                if (!text3.Equals(""))
                {
                    text = text3;
                }
                if (text.Contains("JB") || text.Contains("JC") || text.Contains("(户)") || text.Contains("（户）"))
                {
                    text = text2 + text;
                    break;
                }
                if (dii < 0.05)
                {
                    text = text2 + address;
                    break;
                }
                if (!text.Equals(""))
                {
                    text = text2 + address + text;
                    break;
                }
                if (num2 >= num - 1.0)
                {
                    text = text2;
                }
            }
            list.Add(text);
            return list;
        }

        public void UpdateJzxSm(List<Jzx> jzxList)
        {
            foreach (Jzx current in jzxList)
            {
                this.RepairJzxSm(current);
            }
            foreach (Jzx current in jzxList)
            {
                List<string> jzxSm = current.getJzxSm();
                for (int i = 0; i < jzxSm.Count; i++)
                {
                    string text = jzxSm[i];
                    if (!text.EndsWith("_") && !text.EndsWith("partition") && !text.Contains("宅基地"))
                    {
                        if (i != 0)
                        {
                            for (int j = i - 1; j >= 0; j--)
                            {
                                string text2 = jzxSm[j];
                                if (!text2.EndsWith("_"))
                                {
                                    break;
                                }
                                jzxSm[j] += text.Split(new char[]
								{
									'_'
								})[1];
                            }
                        }
                        for (int j = i + 1; j < jzxSm.Count; j++)
                        {
                            string text2 = jzxSm[j];
                            if (!text2.EndsWith("_"))
                            {
                                break;
                            }
                            jzxSm[j] += text.Split(new char[]
							{
								'_'
							})[1];
                        }
                    }
                }
                if (jzxSm.Count > 1 && this.SelectJzxSmCount("partition", jzxSm) == 5)
                {
                    string text = jzxSm[jzxSm.Count - 2];
                    if (!text.EndsWith("_") && !text.Contains("宅基地") && !text.Contains("空地") && !text.Contains("巷道"))
                    {
                        for (int i = 0; i < jzxSm.Count; i++)
                        {
                            string text3 = jzxSm[i];
                            if (!text3.EndsWith("_") || text3.Equals("partition"))
                            {
                                break;
                            }
                            jzxSm[i] = text3 + text.Split(new char[]
							{
								'_'
							})[1];
                        }
                    }
                    else if (text.EndsWith("_") && !jzxSm[0].EndsWith("_"))
                    {
                        string str = jzxSm[0].Split(this.jzxCh)[2];
                        for (int k = jzxSm.Count - 2; k >= 0; k--)
                        {
                            text = jzxSm[k];
                            if (!text.EndsWith("_"))
                            {
                                break;
                            }
                            jzxSm[k] = text + str;
                        }
                    }
                }
                for (int i = 0; i < jzxSm.Count; i++)
                {
                    string text = jzxSm[i];
                    if (!text.EndsWith("_") && !text.EndsWith("partition") && text.Contains("巷道"))
                    {
                        if (i != 0)
                        {
                            for (int j = i - 1; j >= 0; j--)
                            {
                                string text2 = jzxSm[j];
                                if (!text2.EndsWith("空地"))
                                {
                                    break;
                                }
                                jzxSm[j] = jzxSm[j].Split(new char[]
								{
									'_'
								})[0] + "_" + text.Split(new char[]
								{
									'_'
								})[1];
                            }
                        }
                        for (int j = i + 1; j < jzxSm.Count; j++)
                        {
                            string text2 = jzxSm[j];
                            if (!text2.EndsWith("空地"))
                            {
                                break;
                            }
                            jzxSm[j] = jzxSm[j].Split(new char[]
							{
								'_'
							})[0] + '_' + text.Split(new char[]
							{
								'_'
							})[1];
                        }
                    }
                }
            }
        }

        public int SelectJzxSmCount(string sm, List<string> jzxSm)
        {
            int num = 0;
            foreach (string current in jzxSm)
            {
                if (current.Equals(sm))
                {
                    num++;
                }
            }
            return num;
        }

        public List<string> CheckJzx(List<Jzx> jzxList)
        {
            List<string> list = new List<string>();
            foreach (Jzx current in jzxList)
            {
                List<string> jzxSm = current.getJzxSm();
                for (int i = 0; i < jzxSm.Count; i++)
                {
                    string text = jzxSm[i];
                    if (text.EndsWith("_"))
                    {
                        list.Add(current.getCbfbm() + "，界址线说明不全：" + text);
                    }
                }
                if (this.findJzDh(jzxSm, "J1-") == -1)
                {
                    list.Add(current.getCbfbm() + "，界址线说明没有J1");
                }
            }
            return list;
        }

        public int findJzDh(List<string> jzxSm, string jzxDh)
        {
            int result;
            for (int i = 0; i < jzxSm.Count; i++)
            {
                string text = jzxSm[i];
                if (text.StartsWith(jzxDh) || text.Contains(jzxDh))
                {
                    result = i;
                    return result;
                }
            }
            result = -1;
            return result;
        }

        public void Integration(List<Jzx> jzxList)
        {
            this.MergeJzx(jzxList);
        }

        public void MergeJzx(List<Jzx> jzxList)
        {
            foreach (Jzx current in jzxList)
            {
                this.MergeJzx(current);
            }
        }

        public void MergeJzx(Jzx jzx)
        {
            List<string> list = new List<string>();
            string[] array = new string[3];
            List<string> jzxSm = jzx.getJzxSm();
            for (int i = 0; i < jzxSm.Count; i++)
            {
                string text = jzxSm[i];
                string[] array2 = text.Split(this.jzxCh);
                if (array2.Length != 3 && !text.Equals("partition"))
                {
                    return;
                }
                if (text.Equals("partition") || (array[2] != null && !array[2].Equals(array2[2])))
                {
                    array[1] = jzxSm[i - 1].Split(this.jzxCh)[1];
                    list.Add(string.Concat(new string[]
					{
						array[0],
						"-",
						array[1],
						"_",
						array[2]
					}));
                    array = new string[3];
                    if (!text.Equals("partition"))
                    {
                        array[0] = array2[0];
                        array[1] = array2[1];
                        array[2] = array2[2];
                    }
                }
                else if (array[0] == null)
                {
                    array[0] = array2[0];
                    array[1] = array2[1];
                    array[2] = array2[2];
                }
            }
            jzx.setJzxSm(list);
        }

        public void RepairJzxSm(Jzx jzx)
        {
            List<string> list = new List<string>();
            List<string> jzxSm = jzx.getJzxSm();
            int num = this.findJzDh(jzxSm, "J1-");
            for (int i = num; i < jzxSm.Count; i++)
            {
                list.Add(jzxSm[i]);
            }
            for (int i = 0; i < num; i++)
            {
                list.Add(jzxSm[i]);
            }
            jzx.setJzxSm(list);
        }

        public List<Jzx> SetJzxBs()
        {
            Dictionary<string, string> dictionary = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\行政代码.xls", 0);
            List<Jzx> list = new List<Jzx>();
            Editor editor = this.utils.GetEditor();
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(8, "JZD"),
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding=("\n请选择要生成标示表的JZD图层的权属线");
            PromptSelectionResult selection = editor.GetSelection(promptSelectionOptions, selectionFilter);
            List<Jzx> result;
            if (selection.Status == PromptStatus.Cancel)
            {
                result = list;
            }
            else
            {
                DBObjectCollection entityCollection = this.utils.GetEntityCollection(selection);
                if (entityCollection.Count == 0)
                {
                    editor.WriteMessage("\n你没有选择对象");
                    result = list;
                }
                else
                {
                    Database workingDatabase = HostApplicationServices.WorkingDatabase;
                    using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                    {
                        foreach (DBObject dBObject in entityCollection)
                        {
                            if (dBObject is Polyline)
                            {
                                Jzx jzx = new Jzx("");
                                if (dBObject is Polyline)
                                {
                                    Polyline polyline = (Polyline)transaction.GetObject(dBObject.ObjectId, OpenMode.ForWrite, true);
                                    ResultBuffer resultBuffer = polyline.XData;
                                    if (!(resultBuffer == null))
                                    {
                                        resultBuffer = polyline.GetXDataForApplication("SOUTH");
                                        TypedValue[] array2 = resultBuffer.AsArray();
                                        if (array2 != null && array2.Length >= 5)
                                        {
                                            if (!polyline.Closed)
                                            {
                                                editor.WriteMessage(array2[2].Value.ToString() + "，多段线没有闭合\n");
                                            }
                                            else
                                            {
                                                jzx.setCbfbm(array2[2].Value.ToString());
                                                jzx.setCbfmc(array2[3].Value.ToString());
                                                Point3dCollection point3dCollection = new Point3dCollection();
                                                polyline.GetStretchPoints(point3dCollection);
                                                point3dCollection = this.utils.RePoint(point3dCollection);
                                                List<string> list2 = new List<string>();
                                                double num;
                                                double num2;
                                                double num3;
                                                double num4;
                                                string text;
                                                for (int i = 0; i < point3dCollection.Count - 1; i++)
                                                {
                                                    num = Math.Round(point3dCollection[i].X, 3);
                                                    num2 = Math.Round(point3dCollection[i].Y, 3);
                                                    num3 = Math.Round(point3dCollection[i + 1].X, 3);
                                                    num4 = Math.Round(point3dCollection[i + 1].Y, 3);
                                                    text = this.utils.selectFw_JzxSur(point3dCollection[i], point3dCollection[i + 1], polyline);
                                                    list2.Add(string.Concat(new object[]
													{
														"J",
														i + 1,
														"_J",
														i + 2,
														"_",
														text,
														"_",
														Math.Round(this.utils.CalculateTwoPt(point3dCollection[i], point3dCollection[i + 1]), 6),
														"_",
														num,
														",",
														num2,
														"_",
														num3,
														",",
														num4
													}));
                                                }
                                                num = Math.Round(point3dCollection[point3dCollection.Count - 1].X, 3);
                                                num2 = Math.Round(point3dCollection[point3dCollection.Count - 1].Y, 3);
                                                num3 = Math.Round(point3dCollection[0].X, 3);
                                                num4 = Math.Round(point3dCollection[0].Y, 3);
                                                text = this.utils.selectFw_JzxSur(point3dCollection[point3dCollection.Count - 1], point3dCollection[0], polyline);
                                                list2.Add(string.Concat(new object[]
												{
													"J",
													point3dCollection.Count,
													"_J1_",
													text,
													"_",
													Math.Round(this.utils.CalculateTwoPt(point3dCollection[point3dCollection.Count - 1], point3dCollection[0]), 6),
													"_",
													num,
													",",
													num2,
													"_",
													num3,
													",",
													num4
												}));
                                                jzx.setJzxBs(list2);
                                                list.Add(jzx);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        transaction.Commit();
                    }
                    result = list;
                }
            }
            return result;
        }

        public List<Jzxinfo> JzxToJzxinfo(List<Jzx> jzxList)
        {
            List<Jzxinfo> list = new List<Jzxinfo>();
            char[] separator = new char[]
			{
				'_'
			};
            foreach (Jzx current in jzxList)
            {
                List<string> jzxBs = current.getJzxBs();
                for (int i = 0; i < jzxBs.Count; i++)
                {
                    string text = jzxBs[i];
                    string[] array = text.Split(separator);
                    Jzxinfo jzxinfo = new Jzxinfo();
                    jzxinfo.setBzdh(current.getCbfbm());
                    jzxinfo.setQdh(array[0]);
                    jzxinfo.setZdh(array[1]);
                    jzxinfo.setTsbc(double.Parse(array[4]));
                    jzxinfo.setJxxz("确定界");
                    jzxinfo.setJzxlb(array[2]);
                    jzxinfo.setJzxwz(array[3]);
                    jzxinfo.setQdzb(array[5]);
                    jzxinfo.setZdzb(array[6]);
                    list.Add(jzxinfo);
                }
            }
            return list;
        }

        public List<Qzb> JzxToQzb(List<Jzx> jzxList)
        {
            char[] separator = new char[]
			{
				'_'
			};
            List<Qzb> list = new List<Qzb>();
            foreach (Jzx current in jzxList)
            {
                List<string> jzxSm = current.getJzxSm();
                foreach (string current2 in jzxSm)
                {
                    string[] array = current2.Split(separator);
                    Qzb qzb = new Qzb();
                    qzb.setBzdh(current.getCbfbm());
                    qzb.setQdh(array[0]);
                    qzb.setZdh(array[1]);
                    qzb.setLzdzjr(array[2]);
                    list.Add(qzb);
                }
            }
            return list;
        }

        public List<string> JzxDhChange(List<Jzx> jzxs, Dictionary<string, List<Jzxinfo>> dic)
        {
            List<string> list = new List<string>();
            foreach (Jzx current in jzxs)
            {
                List<string> jzxBs = current.getJzxBs();
                string cbfbm = current.getCbfbm();
                List<Jzxinfo> list2;
                dic.TryGetValue(cbfbm, out list2);
                if (list2 == null)
                {
                    list.Add("mdb中没有找到编码，请查找问题，宗地编码是：" + cbfbm);
                }
                else
                {
                    int count;
                    if (jzxBs.Count != list2.Count)
                    {
                        list.Add("界址线数量与mdb对不上，请查找问题，宗地编码是：" + cbfbm);
                        if (jzxBs.Count > list2.Count)
                        {
                            count = list2.Count;
                        }
                        else
                        {
                            count = jzxBs.Count;
                        }
                    }
                    else
                    {
                        count = jzxBs.Count;
                    }
                    for (int i = 0; i < count; i++)
                    {
                        string text = jzxBs[i];
                        string[] array = text.Split(this.jzxCh);
                        Jzxinfo jzxinfo = list2[i];
                        array[0] = jzxinfo.getQdh();
                        array[1] = jzxinfo.getZdh();
                        jzxBs[i] = Utils.ListToStr(array, "_");
                    }
                }
            }
            return list;
        }

        public Dictionary<string, string> GetjzxBSCorresponding()
        {
            if (JzxService.jzxBSCorresponding == null)
            {
                JzxService.jzxBSCorresponding = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\界址标示表对应位置.xls", 0);
            }
            return JzxService.jzxBSCorresponding;
        }

        public Dictionary<string, string> GetJzxZLDic()
        {
            if (JzxService.jzxZLDic == null)
            {
                JzxService.jzxZLDic = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\界址线种类.xls", 0);
            }
            return JzxService.jzxZLDic;
        }
    }
}
