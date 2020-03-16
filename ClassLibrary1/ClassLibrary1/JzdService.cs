using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using YanBo.office;
using YanBo.ZJDPo;
using YanBo_CG;
using YanBo_CG.ExcelUtils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;

namespace YanBo
{
    internal class JzdService
    {
        private static string FWDCB = "C:\\ZJDTemplete\\项目1\\房屋调查表.xls";

        private static string QJDCB = "C:\\ZJDTemplete\\项目1\\不动产权籍调查表.xls";

        private static Utils utils = new Utils();

        private static ExcelUtils exUtils = new ExcelUtils();

        public static Dictionary<string, string> FWDCB_FW = null;

        public static Dictionary<string, string> FWDCB_PL = null;

        public JzdService()
        {
            if (FWDCB_FW == null)
            {
                FWDCB_FW = ExcelRead.ReadExcelToDic("C:\\ZJDTemplete\\项目1\\房屋包含线.xls", 0);
            }
            if (FWDCB_PL == null)
            {
                FWDCB_PL = ExcelRead.ReadExcelToDic("C:\\ZJDTemplete\\项目1\\界址走向.xls", 0);
            }
        }

        public void createFourAddress(object[] formText)
        {
            Dictionary<string, string> addressDic = ExcelUtils.ReadExcelToDic((string)formText[0], 0);
            Editor editor = utils.GetEditor();
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(8, "JZD"),
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = "\n请选择要生成四至的JZD图层权属线";
            PromptSelectionResult selection = editor.GetSelection(promptSelectionOptions, selectionFilter);
            if (selection.Status == PromptStatus.Cancel)
            {
                DBObjectCollection entityCollection = utils.GetEntityCollection(selection);
                if (entityCollection.Count == 0)
                {
                    editor.WriteMessage("\n你没有选择对象");
                }
                else
                {
                    Database workingDatabase = HostApplicationServices.WorkingDatabase;
                    using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                    {
                        bool cover = (bool)formText[3];
                        foreach (DBObject dBObject in entityCollection)
                        {
                            if (dBObject is Polyline)
                            {
                                ViewTable viewTable = transaction.GetObject(workingDatabase.ViewTableId, 0) as ViewTable;
                                Polyline polyline = (Polyline)transaction.GetObject(dBObject.ObjectId, OpenMode.ForWrite, true);
                                ResultBuffer xData = polyline.XData;
                                if (!(xData == null))
                                {
                                    TypedValue[] array2 = xData.AsArray();
                                    if (array2 != null && array2.Length >= 5)
                                    {
                                        if (array2.Length < 8)
                                        {
                                            array2 = CreateJzd(array2);
                                        }
                                        Zdinfo zdinfo = new Zdinfo();
                                        Point3dCollection point3dCollection = new Point3dCollection();
                                        polyline.GetStretchPoints(point3dCollection);
                                        double[] dou = utils.FindFourPoints(point3dCollection);
                                        Point3dCollection point3dCollection2 = utils.createFourPoint(dou);
                                        Polyline3d polyline3d = utils.CreatePolye(point3dCollection2, false);
                                        int num = utils.PolyDiPoly(point3dCollection, point3dCollection2);
                                        int[] array3 = new int[4];
                                        if (num > 3)
                                        {
                                            int[] array4 = utils.FindPointsDic(point3dCollection);
                                            int[] array5 = utils.FindPointsDic(point3dCollection, point3dCollection2);
                                            int[][] max = new int[][]
											{
												array4,
												array5
											};
                                            int num2 = utils.SelectPc1PC2(max, point3dCollection);
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
                                            array3 = utils.FindPointsDic(point3dCollection, point3dCollection2);
                                        }
                                        double num3 = utils.PointAngle(point3dCollection[array3[0]], point3dCollection[array3[1]]);
                                        double num4 = utils.PointAngle(point3dCollection[array3[0]], point3dCollection[array3[3]]);
                                        double num5 = num3 - num4;
                                        string surAddress = this.GetSurAddress(polyline, addressDic);
                                        for (int i = 0; i < 4; i++)
                                        {
                                            string text = this.CreateSZQS(polyline, point3dCollection, array3, i, 0.05, 0.99, 1);
                                            double num6 = 1.0;
                                            if (text.Equals("") && 1.0 < (double)formText[2])
                                            {
                                                int num7 = 1;
                                                while ((double)num7 <= (double)formText[2])
                                                {
                                                    if (num7 < 4)
                                                    {
                                                        num6 *= 0.98;
                                                    }
                                                    else
                                                    {
                                                        num6 *= 0.96;
                                                    }
                                                    text = this.CreateSZ(polyline, point3dCollection, array3, i, (double)num7, num6, 2);
                                                    if (text.Equals(""))
                                                    {
                                                        text = this.CreateSZ(polyline, point3dCollection, array3, i, (double)num7, num6, 2);
                                                    }
                                                    if (text.Equals(""))
                                                    {
                                                        text = this.CreateSZQS(polyline, point3dCollection, array3, i, (double)num7, num6, 2);
                                                        if (num7 < 5)
                                                        {
                                                            if (!text.Equals(""))
                                                            {
                                                                text = "巷道";
                                                            }
                                                        }
                                                        else if (!text.Equals(""))
                                                        {
                                                            text = "空地";
                                                        }
                                                    }
                                                    if (!text.Equals(""))
                                                    {
                                                        break;
                                                    }
                                                    num7++;
                                                }
                                            }
                                            if (surAddress != null)
                                            {
                                                if (!text.Contains("宅基地"))
                                                {
                                                    text = surAddress + text;
                                                }
                                                if (num5 < 0.0 && num3 < 40.0)
                                                {
                                                    switch (i)
                                                    {
                                                        case 0:
                                                            zdinfo.setBz(text);
                                                            break;
                                                        case 1:
                                                            zdinfo.setDz(text);
                                                            break;
                                                        case 2:
                                                            zdinfo.setNz(text);
                                                            break;
                                                        case 3:
                                                            zdinfo.setXz(text);
                                                            break;
                                                    }
                                                }
                                                else
                                                {
                                                    switch (i)
                                                    {
                                                        case 0:
                                                            zdinfo.setDz(text);
                                                            break;
                                                        case 1:
                                                            zdinfo.setNz(text);
                                                            break;
                                                        case 2:
                                                            zdinfo.setXz(text);
                                                            break;
                                                        case 3:
                                                            zdinfo.setBz(text);
                                                            break;
                                                    }
                                                }
                                            }
                                        }
                                        this.SetZd(array2, zdinfo, cover);
                                        polyline.XData = (new ResultBuffer(array2));
                                    }
                                }
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
        }

        public string GetSurAddress(Polyline poly, Dictionary<string, string> addressDic)
        {
            TypedValue[] surface = this.GetSurface(poly);
            string result;
            if (surface == null)
            {
                result = "";
            }
            else
            {
                string text = surface[2].Value.ToString();
                if (text.Length != 19)
                {
                    throw new Exception("宗地编码不是19位：" + text);
                }
                string key = text.Substring(0, 16).Replace("JC", "").Replace("JB", "");
                string key2 = text.Substring(0, 9);
                string text2;
                addressDic.TryGetValue(key2, out text2);
                string text3;
                addressDic.TryGetValue(key, out text3);
                if (text2 == null || text3 == null)
                {
                    throw new Exception("行政代码表没有这个编码：" + text.Substring(0, 9) + "，或者" + text.Substring(0, 16).Replace("JC", "").Replace("JB", ""));
                }
                result = text3.Replace(text2, "");
            }
            return result;
        }

        public TypedValue[] GetSurface(Polyline poly)
        {
            ResultBuffer resultBuffer = poly.XData;
            TypedValue[] result;
            if (resultBuffer == null)
            {
                result = null;
            }
            else
            {
                resultBuffer = poly.GetXDataForApplication("SOUTH");
                if (resultBuffer != null)
                {
                    TypedValue[] array = resultBuffer.AsArray();
                    if (array.Length == 5)
                    {
                        result = array;
                        return result;
                    }
                }
                result = null;
            }
            return result;
        }

        private string CreateSZ(Polyline poly, Point3dCollection points, int[] pointIndex, int pointH, double di, double sc, int ind)
        {
            Polyline polyline = new Polyline();
            Point3dCollection point3dCollection = new Point3dCollection();
            if (pointIndex[pointH + 1] > pointIndex[pointH])
            {
                for (int i = pointIndex[pointH]; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            else
            {
                for (int i = pointIndex[pointH]; i < points.Count; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
                for (int i = 0; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            Point3dCollection point3dCollection2 = new Point3dCollection();
            for (int j = 0; j < point3dCollection.Count; j++)
            {
                point3dCollection2.Add(point3dCollection[j]);
            }
            for (int j = 0; j < point3dCollection2.Count; j++)
            {
                polyline.AddVertexAt(j, new Point2d(point3dCollection2[j].X, point3dCollection2[j].Y), 0.0, 0.0, 0.0);
            }
            DBObjectCollection offsetCurves = polyline.GetOffsetCurves(0.05);
            point3dCollection = new Point3dCollection();
            foreach (DBObject dBObject in offsetCurves)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection3 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection3);
                    for (int k = 0; k < point3dCollection3.Count; k++)
                    {
                        point3dCollection.Add(point3dCollection3[k]);
                    }
                }
            }
            DBObjectCollection offsetCurves2 = polyline.GetOffsetCurves(-di);
            foreach (DBObject dBObject in offsetCurves2)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection3 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection3);
                    for (int k = point3dCollection3.Count - 1; k >= 0; k--)
                    {
                        point3dCollection.Add(point3dCollection3[k]);
                    }
                }
            }
            Polyline3d poly2 = utils.CreatePolye(point3dCollection, true);
            if (ind == 1)
            {
            }
            utils.Scale(poly2, utils.GetZpoint(point3dCollection[point3dCollection.Count / 2 - 1], point3dCollection[point3dCollection.Count - 1]), sc);
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(poly2);
            string result;
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline poly3 = (Polyline)dBObject;
                    string surface = utils.GetSurface(poly3);
                    if (surface != "")
                    {
                        result = surface;
                        return result;
                    }
                }
                if (dBObject is BlockReference)
                {
                    BlockReference block = (BlockReference)dBObject;
                    string surface = utils.GetSurface(block);
                    if (surface != "")
                    {
                        result = surface;
                        return result;
                    }
                }
                if (dBObject is DBText)
                {
                    DBText dBText = (DBText)dBObject;
                    if (dBText.TextString.Equals("空地"))
                    {
                        result = "空地";
                        return result;
                    }
                    if (dBText.TextString.Equals("空坝"))
                    {
                        result = "空坝";
                        return result;
                    }
                }
            }
            result = "";
            return result;
        }

        private bool RePoint(Point3d point, Point3dCollection newNorthPoints)
        {
            bool result;
            for (int i = 0; i < newNorthPoints.Count; i++)
            {
                if (Math.Abs(point.X - newNorthPoints[i].X) < 0.1 && Math.Abs(point.Y - newNorthPoints[i].Y) < 0.1)
                {
                    result = true;
                    return result;
                }
            }
            result = false;
            return result;
        }

        private string CreateSZQS(Polyline poly, Point3dCollection points, int[] pointIndex, int pointH, double di, double sc, int ind)
        {
            Polyline polyline = new Polyline();
            Point3dCollection point3dCollection = new Point3dCollection();
            if (pointIndex[pointH + 1] > pointIndex[pointH])
            {
                for (int i = pointIndex[pointH]; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            else
            {
                for (int i = pointIndex[pointH]; i < points.Count; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
                for (int i = 0; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            for (int j = 0; j < point3dCollection.Count; j++)
            {
                polyline.AddVertexAt(j, new Point2d(point3dCollection[j].X, point3dCollection[j].Y), 0.0, 0.0, 0.0);
            }
            polyline.AddVertexAt(0, new Point2d(point3dCollection[0].X, point3dCollection[0].Y), 0.0, 0.0, 0.0);
            DBObjectCollection offsetCurves = polyline.GetOffsetCurves(0.05);
            point3dCollection = new Point3dCollection();
            foreach (DBObject dBObject in offsetCurves)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int k = 0; k < point3dCollection2.Count; k++)
                    {
                        point3dCollection.Add(point3dCollection2[k]);
                    }
                }
            }
            DBObjectCollection offsetCurves2 = polyline.GetOffsetCurves(-di);
            foreach (DBObject dBObject in offsetCurves2)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int k = point3dCollection2.Count - 1; k >= 0; k--)
                    {
                        point3dCollection.Add(point3dCollection2[k]);
                    }
                }
            }
            Polyline3d poly2 = utils.CreatePolye(point3dCollection, true);
            utils.Scale(poly2, utils.GetZpoint(point3dCollection[point3dCollection.Count / 2 - 1], point3dCollection[point3dCollection.Count - 1]), sc);
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(poly2);
            string result;


            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    string text = "";
                    Polyline poly3 = (Polyline)dBObject;
                    object[] surface = utils.GetSurface(poly3, poly);
                    TypedValue[] array = null;
                    if (surface != null)
                    {
                        array = (TypedValue[])surface[0];
                    }
                    if (array != null && array.Length > 3)
                    {
                        text = array[3].Value.ToString() + "宅基地";
                    }
                    if (text != "")
                    {
                        result = text;
                        return result;
                    }
                }
            }
            result = "";
            return result;
        }

        public DBObjectCollection SelectJzd(Polyline poly, Point3dCollection points, int[] pointIndex, int pointH, double di, double sc, int ind)
        {
            Polyline polyline = new Polyline();
            Point3dCollection point3dCollection = new Point3dCollection();
            if (pointIndex[pointH + 1] > pointIndex[pointH])
            {
                for (int i = pointIndex[pointH]; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            else
            {
                for (int i = pointIndex[pointH]; i < points.Count; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
                for (int i = 0; i <= pointIndex[pointH + 1]; i++)
                {
                    if (!this.RePoint(new Point3d(points[i].X, points[i].Y, 0.0), point3dCollection))
                    {
                        point3dCollection.Add(new Point3d(points[i].X, points[i].Y, 0.0));
                    }
                }
            }
            for (int j = 0; j < point3dCollection.Count; j++)
            {
                polyline.AddVertexAt(j, new Point2d(point3dCollection[j].X, point3dCollection[j].Y), 0.0, 0.0, 0.0);
            }
            polyline.AddVertexAt(0, new Point2d(point3dCollection[0].X, point3dCollection[0].Y), 0.0, 0.0, 0.0);
            DBObjectCollection offsetCurves = polyline.GetOffsetCurves(0.05);
            point3dCollection = new Point3dCollection();
            foreach (DBObject dBObject in offsetCurves)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int k = 0; k < point3dCollection2.Count; k++)
                    {
                        point3dCollection.Add(point3dCollection2[k]);
                    }
                }
            }
            DBObjectCollection offsetCurves2 = polyline.GetOffsetCurves(-di);
            foreach (DBObject dBObject in offsetCurves2)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int k = point3dCollection2.Count - 1; k >= 0; k--)
                    {
                        point3dCollection.Add(point3dCollection2[k]);
                    }
                }
            }
            Point3dCollection pc = CADUtils.ScalePoints(point3dCollection, utils.GetZpoint(point3dCollection[point3dCollection.Count / 2 - 1], point3dCollection[point3dCollection.Count - 1]), sc);
            return utils.SelectJzd(pc);
        }

        public Zdinfo GetZdinfo(Polyline poly)
        {
            Zdinfo zdinfo = new Zdinfo();
            ResultBuffer xData = poly.XData;
            Zdinfo result;
            if (xData == null)
            {
                result = zdinfo;
            }
            else
            {
                TypedValue[] array = xData.AsArray();
                if (array != null && array.Length >= 69)
                {
                    int i = 0;
                    while (i < array.Length)
                    {
                        string text = ((TypedValue)array.GetValue(i)).Value.ToString();
                        if (text != null)
                        {
                            if (!(text == "BEIZHI"))
                            {
                                if (!(text == "DONGZHI"))
                                {
                                    if (!(text == "NANZHI"))
                                    {
                                        if (text == "XIZHI")
                                        {
                                            zdinfo.setXz(array[i + 1].Value.ToString());
                                        }
                                    }
                                    else
                                    {
                                        zdinfo.setNz(array[i + 1].Value.ToString());
                                    }
                                }
                                else
                                {
                                    zdinfo.setDz(array[i + 1].Value.ToString());
                                }
                            }
                            else
                            {
                                zdinfo.setBz(array[i + 1].Value.ToString());
                            }
                        }
                    IL_11C:
                        i++;
                        continue;
                        goto IL_11C;
                    }
                }
                result = zdinfo;
            }
            return result;
        }

        public void CreateZDT_SZ()
        {
            Utils.CreateLayer("宗地四至");
            Editor editor = utils.GetEditor();
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(8, "JZD")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = ("\n请选择要生成四至的JZD图层权属线");
            PromptSelectionResult selection = editor.GetSelection(promptSelectionOptions, selectionFilter);
            if (selection.Status != PromptStatus.Cancel)
            {
                DBObjectCollection entityCollection = utils.GetEntityCollection(selection);
                if (entityCollection.Count == 0)
                {
                    editor.WriteMessage("\n你没有选择对象");
                }
                else
                {
                    Database workingDatabase = HostApplicationServices.WorkingDatabase;
                    ObjectId textStyle = utils.CreateSytle1("细等线体", 1.0);
                    using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                    {
                        foreach (DBObject dBObject in entityCollection)
                        {
                            if (dBObject is Polyline)
                            {
                                Polyline polyline = dBObject as Polyline;
                                if (polyline.Closed)
                                {
                                    string layerName = "四至注记";
                                    //Utils.CreateLayer(layerName);

                                    Zdinfo zdinfo = this.GetZdinfo(polyline);
                                    Point3dCollection jzdFourExPoints = this.GetJzdFourExPoints(polyline);


                                    //ObjectId textStyle = utils.CreateSytle1("细等线体", 1.0);
                                    DBText dBText = utils.CreateDBText(GetSZ(zdinfo.getBz()), jzdFourExPoints[0], 1.2, textStyle);
                                    dBText.HorizontalMode = TextHorizontalMode.TextCenter;
                                    dBText.VerticalMode = TextVerticalMode.TextBottom;
                                    dBText.AlignmentPoint = dBText.Position;
                                    dBText.Layer = (layerName);
                                    utils.AddDBText(dBText);

                                    string dz = zdinfo.getDz();
                                    DBText dBText2 = utils.CreateDBText(GetSZ(zdinfo.getDz()), jzdFourExPoints[1], 1.2, textStyle);
                                    dBText2.Layer = (layerName);

                                    dBText2.HorizontalMode = TextHorizontalMode.TextCenter;
                                    dBText2.VerticalMode = TextVerticalMode.TextBottom;
                                    dBText2.AlignmentPoint = dBText2.Position;

                                    utils.AddDBText(dBText2);

                                    string nz = zdinfo.getNz();
                                    DBText dBText3 = utils.CreateDBText(GetSZ(zdinfo.getNz()), jzdFourExPoints[2], 1.2, textStyle);
                                    dBText3.Layer = (layerName);

                                    dBText3.HorizontalMode = TextHorizontalMode.TextCenter;
                                    dBText3.VerticalMode = TextVerticalMode.TextBottom;
                                    dBText3.AlignmentPoint = dBText3.Position;


                                    utils.AddDBText(dBText3);

                                    string xz = zdinfo.getXz();
                                    DBText dBText4 = utils.CreateDBText(GetSZ(zdinfo.getXz()), jzdFourExPoints[3], 1.2, textStyle);

                                    dBText4.HorizontalMode = TextHorizontalMode.TextCenter;
                                    dBText4.VerticalMode = TextVerticalMode.TextBottom;
                                    dBText4.AlignmentPoint = dBText4.Position;

                                    dBText4.Layer = (layerName);


                                   
                                    utils.AddDBText(dBText4);
                                }
                            }
                        }
                        transaction.Commit();
                    }
                }
            }
        }

        public string GetSZ(string str)
        {
            if (str != null)
            {
                if (str.Contains(",") || str.Contains("，"))
                {
                    char[] ch = new char[]{
                        ',','，'
                    };
                    string[] array = str.Split(ch);
                    if (array.Length > 1)
                    {
                        return array[1].Replace("邻", "");
                    }
                }
                else 
                {
                    if (str.Length > 4)
                    {
                        str = str.Substring(1, str.Length - 3)+"房子";
                        return "";
                    }
                   
                   
                }
                
            } 
                return str;
            
        }
        public void SetZd(TypedValue[] tv, Zdinfo zdinfo, bool cover)
        {
            if (tv.Length >= 69)
            {
                for (int i = 0; i < tv.Length; i++)
                {
                    TypedValue typedValue = (TypedValue)tv.GetValue(i);
                    if (typedValue.Value != null)
                    {
                        string text = typedValue.Value.ToString();
                        if (text != null)
                        {
                            if (!(text == "BEIZHI"))
                            {
                                if (!(text == "DONGZHI"))
                                {
                                    if (!(text == "NANZHI"))
                                    {
                                        if (text == "XIZHI")
                                        {
                                            if (cover)
                                            {
                                                tv[i + 1] = new TypedValue(1000, zdinfo.getXz());
                                            }
                                            else if (tv[i + 1].Value == null || tv[i + 1].Value.ToString().Equals(""))
                                            {
                                                tv[i + 1] = new TypedValue(1000, zdinfo.getXz());
                                            }
                                        }
                                    }
                                    else if (cover)
                                    {
                                        tv[i + 1] = new TypedValue(1000, zdinfo.getNz());
                                    }
                                    else if (tv[i + 1].Value == null || tv[i + 1].Value.ToString().Equals(""))
                                    {
                                        tv[i + 1] = new TypedValue(1000, zdinfo.getNz());
                                    }
                                }
                                else
                                {
                                    if (cover)
                                    {
                                        tv[i + 1] = new TypedValue(1000, zdinfo.getDz());
                                    }
                                    if (tv[i + 1].Value == null || tv[i + 1].Value.ToString().Equals(""))
                                    {
                                        tv[i + 1] = new TypedValue(1000, zdinfo.getDz());
                                    }
                                }
                            }
                            else if (cover)
                            {
                                tv[i + 1] = new TypedValue(1000, zdinfo.getBz());
                            }
                            else if (tv[i + 1].Value == null || tv[i + 1].Value.ToString().Equals(""))
                            {
                                tv[i + 1] = new TypedValue(1000, zdinfo.getBz());
                            }
                        }
                    }
                }
            }
        }

        public Point3dCollection GetJzdFourExPoints(Polyline poly)
        {
            Point3dCollection jzdFourMiPoints = this.GetJzdFourMiPoints(poly);
            Point3dCollection point3dCollection = new Point3dCollection();
            Point3d point3d = new Point3d(jzdFourMiPoints[0].X, jzdFourMiPoints[0].Y + 7.0, 0.0);
            point3dCollection.Add(point3d);
            point3d = new Point3d(jzdFourMiPoints[1].X + 5.0, jzdFourMiPoints[1].Y, 0.0);
            point3dCollection.Add(point3d);
            point3d = new Point3d(jzdFourMiPoints[2].X, jzdFourMiPoints[2].Y - 10.0, 0.0);
            point3dCollection.Add(point3d);
            point3d = new Point3d(jzdFourMiPoints[3].X - 10.0, jzdFourMiPoints[3].Y, 0.0);
            point3dCollection.Add(point3d);
            return point3dCollection;
        }

        public Point3dCollection GetJzdFourMiPoints(Polyline poly)
        {
            Point3dCollection jzdFourPoints = this.GetJzdFourPoints(poly);
            Point3dCollection point3dCollection = new Point3dCollection();
            point3dCollection.Add(utils.GetZpoint(jzdFourPoints[0], jzdFourPoints[1]));
            point3dCollection.Add(utils.GetZpoint(jzdFourPoints[1], jzdFourPoints[2]));
            point3dCollection.Add(utils.GetZpoint(jzdFourPoints[2], jzdFourPoints[3]));
            point3dCollection.Add(utils.GetZpoint(jzdFourPoints[3], jzdFourPoints[0]));
            return point3dCollection;
        }

        public  Point3dCollection GetJzdFourPoints(Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            double[] dou = utils.FindFourPoints(point3dCollection);
            Point3dCollection point3dCollection2 = utils.createFourPoint(dou);
            Polyline3d polyline3d = utils.CreatePolye(point3dCollection2, false);
            int num = utils.PolyDiPoly(point3dCollection, point3dCollection2);
            int[] array = new int[4];
            if (num > 3)
            {
                int[] array2 = utils.FindPointsDic(point3dCollection);
                int[] array3 = utils.FindPointsDic(point3dCollection, point3dCollection2);
                int[][] max = new int[][]
				{
					array2,
					array3
				};
                int num2 = utils.SelectPc1PC2(max, point3dCollection);
                if (num2 == 0)
                {
                    array = array2;
                }
                else if (num2 == 1)
                {
                    array = array3;
                }
            }
            else
            {
                array = utils.FindPointsDic(point3dCollection, point3dCollection2);
            }
            double num3 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[1]]);
            double num4 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[3]]);
            double num5 = num3 - num4;
            Point3dCollection point3dCollection3 = new Point3dCollection();
            if (num5 < 0.0 && num3 < 40.0)
            {
                point3dCollection3.Add(point3dCollection[array[0]]);
                point3dCollection3.Add(point3dCollection[array[1]]);
                point3dCollection3.Add(point3dCollection[array[2]]);
                point3dCollection3.Add(point3dCollection[array[3]]);
            }
            else
            {
                point3dCollection3.Add(point3dCollection[array[3]]);
                point3dCollection3.Add(point3dCollection[array[0]]);
                point3dCollection3.Add(point3dCollection[array[1]]);
                point3dCollection3.Add(point3dCollection[array[2]]);
            }
            return point3dCollection3;
        }

        public static void CreateJzd(Polyline poly)
        {
            ResultBuffer xData = poly.XData;
            TypedValue[] values = xData.AsArray();
            CreateJzd(values);
        }

        public static TypedValue[] CreateJzd(TypedValue[] values)
        {
            TypedValue[] array = new TypedValue[69];
            for (int i = 0; i < values.Length; i++)
            {
                TypedValue typedValue = values[i];
                array.SetValue(typedValue, i);
            }
            array.SetValue(new TypedValue(1001, "SJZGBM"), 7);
            array.SetValue(new TypedValue(1000, null), 8);
            array.SetValue(new TypedValue(1001, "FRDBXM"), 9);
            array.SetValue(new TypedValue(1000, null), 10);
            array.SetValue(new TypedValue(1001, "FRDBZMS"), 11);
            array.SetValue(new TypedValue(1000, null), 12);
            array.SetValue(new TypedValue(1001, "FRDBDH"), 13);
            array.SetValue(new TypedValue(1000, null), 14);
            array.SetValue(new TypedValue(1001, "DLRXM"), 15);
            array.SetValue(new TypedValue(1000, null), 16);
            array.SetValue(new TypedValue(1001, "DLRSFZ"), 17);
            array.SetValue(new TypedValue(1000, null), 18);
            array.SetValue(new TypedValue(1001, "DLRDH"), 19);
            array.SetValue(new TypedValue(1000, null), 20);
            array.SetValue(new TypedValue(1001, "TXDZ"), 21);
            array.SetValue(new TypedValue(1000, null), 22);
            array.SetValue(new TypedValue(1001, "TDZL"), 23);
            array.SetValue(new TypedValue(1000, null), 24);
            array.SetValue(new TypedValue(1001, "DONGZHI"), 25);
            array.SetValue(new TypedValue(1000, null), 26);
            array.SetValue(new TypedValue(1001, "NANZHI"), 27);
            array.SetValue(new TypedValue(1000, null), 28);
            array.SetValue(new TypedValue(1001, "XIZHI"), 29);
            array.SetValue(new TypedValue(1000, null), 30);
            array.SetValue(new TypedValue(1001, "BEIZHI"), 31);
            array.SetValue(new TypedValue(1000, null), 32);
            array.SetValue(new TypedValue(1001, "QSLYZM"), 33);
            array.SetValue(new TypedValue(1000, null), 34);
            array.SetValue(new TypedValue(1001, "PZTDYT"), 35);
            array.SetValue(new TypedValue(1000, null), 36);
            array.SetValue(new TypedValue(1001, "TDSYZ"), 37);
            array.SetValue(new TypedValue(1000, null), 38);
            array.SetValue(new TypedValue(1001, "SBJZWQS"), 39);
            array.SetValue(new TypedValue(1000, null), 40);
            array.SetValue(new TypedValue(1001, "YBDJH"), 41);
            array.SetValue(new TypedValue(1000, null), 42);
            array.SetValue(new TypedValue(1001, "TDZH"), 43);
            array.SetValue(new TypedValue(1000, null), 44);
            array.SetValue(new TypedValue(1001, "SHRQ"), 45);
            array.SetValue(new TypedValue(1000, null), 46);
            array.SetValue(new TypedValue(1001, "DJRQ"), 47);
            array.SetValue(new TypedValue(1000, null), 48);
            array.SetValue(new TypedValue(1001, "ZZRQ"), 49);
            array.SetValue(new TypedValue(1000, null), 50);
            array.SetValue(new TypedValue(1001, "DWXZ"), 51);
            array.SetValue(new TypedValue(1000, null), 52);
            array.SetValue(new TypedValue(1001, "QSXZ"), 53);
            array.SetValue(new TypedValue(1000, null), 54);
            array.SetValue(new TypedValue(1001, "SYQLX"), 55);
            array.SetValue(new TypedValue(1000, null), 56);
            array.SetValue(new TypedValue(1001, "TDDJ"), 57);
            array.SetValue(new TypedValue(1000, null), 58);
            array.SetValue(new TypedValue(1001, "MPH"), 59);
            array.SetValue(new TypedValue(1000, null), 60);
            array.SetValue(new TypedValue(1001, "TUFU"), 61);
            array.SetValue(new TypedValue(1000, null), 62);
            array.SetValue(new TypedValue(1001, "JZMJ"), 63);
            array.SetValue(new TypedValue(1000, null), 64);
            array.SetValue(new TypedValue(1001, "BDDJ"), 65);
            array.SetValue(new TypedValue(1040, null), 66);
            array.SetValue(new TypedValue(1001, "SBDJ"), 67);
            array.SetValue(new TypedValue(1040, null), 68);
            return array;
        }

        private void CreateJZDTable(IList<Zdinfo> zds, string saveName)
        {
            if (zds != null && zds.Count != 0)
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet;
                if (workbook.NumberOfSheets == 0)
                {
                    sheet = workbook.CreateSheet();
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                for (int i = 0; i < zds.Count; i++)
                {
                    Zdinfo zdinfo = zds[i];
                    IRow row = sheet.CreateRow(i);
                    for (int j = 0; j < 3; j++)
                    {
                        ICell cell = row.CreateCell(j);
                        switch (j)
                        {
                            case 0:
                                cell.SetCellValue(zdinfo.getZdnum());
                                break;
                            case 1:
                                cell.SetCellValue(zdinfo.getQuanli());
                                break;
                            case 2:
                                cell.SetCellValue(this.PCToString(zdinfo.PC));
                                break;
                        }
                    }
                }
                ExcelWrite.Save(workbook, saveName);
            }
        }

        private string PCToString(Point3dCollection point3dCollection)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (Point3d point3d in point3dCollection)
            {
                stringBuilder.Append(string.Concat(new object[]
				{
					point3d.X,
					":",
					point3d.Y,
					","
				}));
            }
            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            return stringBuilder.ToString();
        }

        private static Point3dCollection ToPC(DBObject dbo)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            if (dbo is Polyline2d)
            {
                Polyline2d polyline2d = (Polyline2d)dbo;
                polyline2d.GetStretchPoints(point3dCollection);
            }
            else if (dbo is Polyline)
            {
                Polyline polyline = (Polyline)dbo;
                polyline.GetStretchPoints(point3dCollection);
            }
            return point3dCollection;
        }

        internal void ExcportJZDXLX()
        {
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = ("\n请选择要生成的权属线");
            PromptSelectionResult selection = utils.GetEditor().GetSelection(promptSelectionOptions, selectionFilter);
            DBObjectCollection entityCollection = utils.GetEntityCollection(selection);
            IList<Zdinfo> list = new List<Zdinfo>();
            foreach (DBObject dBObject in entityCollection)
            {
                ResultBuffer xDataForApplication = dBObject.GetXDataForApplication("SOUTH");
                if (xDataForApplication != null)
                {
                    TypedValue[] array2 = xDataForApplication.AsArray();
                    if (array2.Length > 3)
                    {
                        Zdinfo zdinfo = new Zdinfo();
                        zdinfo.setZdnum(array2[2].Value.ToString());
                        zdinfo.setQuanli(array2[3].Value.ToString());
                        zdinfo.PC = ToPC(dBObject);
                        list.Add(zdinfo);
                    }
                }
            }
            this.CreateJZDTable(list, utils.GetDocument().Name.Replace("dwg", "xlsx"));
        }

        public static DBObjectCollection SelectJZD(string tip)
        {
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding = tip;
            PromptSelectionResult selection = utils.GetEditor().GetSelection(promptSelectionOptions, selectionFilter);
            return utils.GetEntityCollection(selection);
        }

        internal void ExportJzdTable(string saveName)
        {
            DBObjectCollection dBObjectCollection = SelectJZD("请选择要生成的权属线\n");
            IList<Zdinfo> list = new List<Zdinfo>();
            foreach (DBObject dbo in dBObjectCollection)
            {
                Zdinfo zdinfo = ToZdinfo(dbo);
                if (zdinfo != null)
                {
                    list.Add(zdinfo);
                }
            }
            this.CreateJZDCGTable(list, saveName);
        }

        public static Zdinfo ToZdinfo(DBObject dbo)
        {
            Zdinfo result;
            if (!(dbo is Polyline))
            {
                result = null;
            }
            else
            {
                ResultBuffer xData = dbo.XData;
                if (xData != null)
                {
                    TypedValue[] array = xData.AsArray();
                    if (array.Length > 10)
                    {
                        if (array.Length < 10)
                        {
                        }
                        Zdinfo zdinfo = new Zdinfo();
                        zdinfo.PC = ToPC(dbo);
                        for (int i = 0; i < array.Length; i++)
                        {
                            string text = array[i].Value.ToString();
                            switch (text)
                            {
                                case "SOUTH":
                                    zdinfo.setZdnum(array[i + 2].Value.ToString());
                                    zdinfo.setQuanli(array[i + 3].Value.ToString());
                                    break;
                                case "JZMJ":
                                    zdinfo.jzmj = array[i + 1].Value.ToString();
                                    break;
                                case "DONGZHI":
                                    zdinfo.setDz(array[i + 1].Value.ToString());
                                    break;
                                case "NANZHI":
                                    zdinfo.setNz(array[i + 1].Value.ToString());
                                    break;
                                case "XIZHI":
                                    zdinfo.setXz(array[i + 1].Value.ToString());
                                    break;
                                case "BEIZHI":
                                    zdinfo.setBz(array[i + 1].Value.ToString());
                                    break;
                            }
                        }
                        if (zdinfo.getZdnum() == null)
                        {
                            result = null;
                            return result;
                        }
                        zdinfo.Polyline = (Polyline)dbo;
                        result = zdinfo;
                        return result;
                    }
                }
                result = null;
            }
            return result;
        }

        private void CreateJZDCGTable(IList<Zdinfo> zds, string saveName)
        {
            string fileName = "C:\\ZJDTemplete\\界址点成果表.xls";
            IWorkbook workbook = ExcelRead.ReadExcel(fileName);
            ISheet sheetAt = workbook.GetSheetAt(0);
            string newValue = DateTime.Now.ToString("yyyy年M月dd日");
            ICell cell = sheetAt.GetRow(81).GetCell(0);
            cell.SetCellValue(cell.StringCellValue.Replace("DateTime", newValue));
            ISheet sheet = null;
            foreach (Zdinfo current in zds)
            {
                try
                {
                    sheet = sheetAt.CopySheet(current.getZdnum(), true);
                }
                catch
                {
                    MessageBox.Show("宗地编码重复：" + current.getZdnum());
                    continue;
                }
                Point3dCollection pC = current.PC;
                if (pC[0].X != pC[pC.Count - 1].X && pC[0].Y != pC[pC.Count - 1].Y)
                {
                    pC.Add(pC[0]);
                }
                int num = pC.Count / 35 + 1;
                if (num > 2)
                {
                    MessageBox.Show("界址点成果表超过两页，此页没有生成：" + current.getZdnum());
                }
                else if (num == 1)
                {
                    for (int i = 83; i <= sheet.LastRowNum; i++)
                    {
                        sheet.RemoveRow(sheet.GetRow(i));
                    }
                }
                else
                {
                    cell = sheet.GetRow(85).GetCell(0);
                    cell.SetCellValue(cell.StringCellValue.Replace("ZDNUM", current.getZdnum()));
                    cell = sheet.GetRow(86).GetCell(0);
                    cell.SetCellValue(cell.StringCellValue.Replace("QUANLI", current.getQuanli()));
                    cell = sheet.GetRow(166).GetCell(0);
                    cell.SetCellValue(cell.StringCellValue.Replace("DateTime", newValue));
                }
                IRow row = sheet.GetRow(1);
                cell = row.GetCell(4);
                cell.SetCellValue(cell.StringCellValue.Replace("1", num + ""));
                cell = sheet.GetRow(2).GetCell(0);
                cell.SetCellValue(cell.StringCellValue.Replace("ZDNUM", current.getZdnum()));
                cell = sheet.GetRow(3).GetCell(0);
                cell.SetCellValue(cell.StringCellValue.Replace("QUANLI", current.getQuanli()));
                cell = sheet.GetRow(4).GetCell(0);
                cell.SetCellValue(cell.StringCellValue.Replace("ZDMJ", current.area));
                cell = sheet.GetRow(5).GetCell(0);
                cell.SetCellValue(cell.StringCellValue.Replace("JZMJ", current.jzmj));
                int num2 = 0;
                for (int i = 0; i < num; i++)
                {
                    int num3 = 9;
                    if (i == 1)
                    {
                        num3 = 90;
                    }
                    int num4;
                    if (i == 0)
                    {
                        if (pC.Count < 35)
                        {
                            num4 = pC.Count;
                        }
                        else
                        {
                            num4 = 35;
                        }
                    }
                    else
                    {
                        num4 = pC.Count - 35;
                    }
                    for (int j = 0; j < num4; j++)
                    {
                        row = sheet.GetRow(num3);
                        num3 += 2;
                        Point3d pt = pC[num2++];
                        for (int k = 0; k < 5; k++)
                        {
                            cell = row.GetCell(k);
                            switch (k)
                            {
                                case 0:
                                    cell.SetCellValue((double)num2);
                                    break;
                                case 1:
                                    if (num2 == pC.Count)
                                    {
                                        cell.SetCellValue("J1");
                                    }
                                    else
                                    {
                                        cell.SetCellValue("J" + num2);
                                    }
                                    break;
                                case 2:
                                    cell.SetCellValue(pt.X);
                                    break;
                                case 3:
                                    cell.SetCellValue(pt.Y);
                                    break;
                                case 4:
                                    if (j < num4 - 1 && j != 36)
                                    {
                                        sheet.GetRow(row.RowNum + 1).GetCell(4).SetCellValue(utils.CalculateTwoPt(pt, pC[num2]));
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
            workbook.SetSelectedTab(0);
            workbook.RemoveSheetAt(0);
            ExcelWrite.Save(workbook, saveName);
        }

        public int JiSuanHang(int a)
        {
            return a * 83 + a;
        }

        public  void ExportFWDCB(string saveDir, string zdxPath)
        {
            DBObjectCollection dbc = SelectJZD("请选择要生成房屋调查表的权属线\n");
            IList<Zdinfo> zds = DBCToZdinfo(dbc);
            IList<ZDX> zdxs = ExcelToZdx(zdxPath);
            SetZDX(zds, GetZdxDic(zdxs));
            ExportFWDCB(zds, saveDir);
        }

        public static Dictionary<string, ZDX> GetZdxDic(IList<ZDX> zdxs)
        {
            Dictionary<string, ZDX> dictionary = new Dictionary<string, ZDX>();
            foreach (ZDX current in zdxs)
            {
                string zDNUM = current.ZDNUM;
                if (zDNUM == null)
                {
                    throw new Exception("新宗地表中有空的宗地编码");
                }
                if (dictionary.ContainsKey(zDNUM))
                {
                    throw new Exception("新宗地表宗地编码重复：" + zDNUM);
                }
                dictionary.Add(zDNUM, current);
            }
            return dictionary;
        }

        public static void SetZDX(IList<Zdinfo> zds, Dictionary<string, ZDX> zdxDic)
        {
            foreach (Zdinfo current in zds)
            {
                string zdnum = current.getZdnum();
                if (zdnum == null || zdnum.Trim().Equals(""))
                {
                    MessageBox.Show("地籍图上有宗地编码是空的：" + zdnum);
                }
                else
                {
                    ZDX zdx;
                    if (!zdxDic.TryGetValue(zdnum, out zdx))
                    {
                        MessageBox.Show("图上宗地编码在  新宗地表中没有找到：" + zdnum);
                        break;
                    }
                    current.zdx = zdx;
                }
            }
        }

        private static IList<ZDX> ExcelToZdx(string zdxPath)
        {
            IList<ZDX> list = new List<ZDX>();
            IWorkbook workbook = ExcelRead.ReadExcel(zdxPath);
            ISheet sheetAt = workbook.GetSheetAt(0);
            for (int i = 1; i <= sheetAt.LastRowNum; i++)
            {
                IRow row = sheetAt.GetRow(i);
                if (row != null)
                {
                    ZDX zDX = new ZDX();
                    for (int j = 0; j < 50; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            cell.SetCellType((NPOI.SS.UserModel.CellType.String));
                            string stringCellValue = cell.StringCellValue;

                               
                                    switch (j)
                                    {
                                        case 0:
                                            zDX.ZDNUM = stringCellValue;
                                            break;
                                        case 2:
                                            zDX.SFZHM = stringCellValue;
                                            break;
                                        case 4:
                                            zDX.LXDH = stringCellValue;
                                            break;
                                        case 18:
                                            zDX.DZWFWDM = stringCellValue;
                                            break;
                                        case 19:
                                            zDX.BDCDYH = stringCellValue;
                                            break;
                                        case 20:
                                            zDX.TXDZ = stringCellValue;
                                            break;
                                        case 21:
                                            zDX.YZBM = stringCellValue;
                                            break;
                                        case 23:
                                            zDX.ZJZL = stringCellValue;
                                            break;
                                        case 25:
                                            zDX.FWXZ = stringCellValue;
                                            break;
                                        case 26:
                                            zDX.YT = stringCellValue;
                                            break;
                                        case 27:
                                            zDX.GHYT = stringCellValue;
                                            break;
                                        case 32:
                                            zDX.QLRLX = stringCellValue;
                                            break;

                                             
                                    }
                              
                            
                        }
                    }
                    if (zDX.ZDNUM != null)
                    {
                        list.Add(zDX);
                    }
                   
                }
                
            }
            return list;
        }

        private  void ExportFWDCB(IList<Zdinfo> zds, string dir)
        {
            foreach (Zdinfo current in zds)
            {
                IList<JMD> jdms = SelectContainsFWX(current);
                IWorkbook workbook = ExportFWDCB(current, jdms);
                if (workbook != null)
                {
                    ExcelWrite.Save(workbook, dir + "\\" + current.getZdnum() + "_房屋调查表.xls");
                }
            }
        }

        private  IList<JMD> SelectContainsFWX(Zdinfo zdinfo)
        {
            IList<JMD> list = new List<JMD>();
            IList<JMD> listOut = new List<JMD>();
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(zdinfo.PC);
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline = zdinfo.Polyline;
                    Polyline polyline2 = (Polyline)dBObject;
                    if (polyline2.Area <= 86.0 && polyline2.Area > 85.0)
                    {
                    }
                    if (!(polyline == polyline2))
                    {
                        Autodesk.AutoCAD.DatabaseServices.Region region = CADUtils.CreateRegion(zdinfo.Polyline);
                        Autodesk.AutoCAD.DatabaseServices.Region region2 = CADUtils.CreateRegion(polyline2);
                        if (region2 != null && region2.Area != 0.0)
                        {
                            double area = region.Area;
                            try
                            {
                                region.BooleanOperation(0, region2);
                                JMD jMD = PLToJMD(polyline2, FWDCB_FW);
                                if (jMD != null)
                                {
                                    if (region.Area < area + 0.3)
                                    {


                                        list.Add(jMD);

                                    }
                                    else
                                    {
                                        listOut.Add(jMD);
                                    }
                                }


                            }
                            catch
                            {
                            }
                        }
                    }
                }
            }
            //查看与外面是否是共墙
             SetJMD_QTGS(list, listOut);
            return list;
        }

        private  void SetJMD_QTGS(IList<JMD> list, IList<JMD> listOut)
        {
            string gy = "共有";
            string zy = "自有";
            if (listOut.Count == 0)
            {
                
                foreach (JMD jmd in list)
                {
                    SetQTGS(jmd, zy);
                    
                }
                
                return;
            }
            IList<JMD> jmds;
            foreach (JMD jmd in list)
            {
               
                DBObjectCollection[] dbcs = CADUtils.GetScalePoints(jmd.PL);
               
                jmds = PLToJMD(dbcs[0],FWDCB_FW);
                if (JMDInDBC(listOut, jmds))
                {
                    jmd.QTGSB = gy;
                }
                else
                {
                    jmd.QTGSB = zy;
                }
                
               

                 jmds = PLToJMD(dbcs[1], FWDCB_FW);

                 if (JMDInDBC(listOut, jmds))
                 {
                     jmd.QTGSD = gy;
                 }
                 else
                 {
                     jmd.QTGSD = zy;
                 }
                

                
                jmds = PLToJMD(dbcs[2], FWDCB_FW);

                if (JMDInDBC(listOut, jmds))
                {
                    jmd.QTGSN = gy;
                }
                else
                {
                    jmd.QTGSN = zy;
                }

               
                jmds = PLToJMD(dbcs[3], FWDCB_FW);
                if (JMDInDBC(listOut, jmds))
                {
                    jmd.QTGSX = gy;
                }
                else
                {
                    jmd.QTGSX = zy;
                }

                
            }

        }

        private bool JMDInDBC(IList<JMD> listOut, IList<JMD> jmds)
        {
            if (jmds.Count == 0)
            {
                return false;
            }
            foreach (JMD jmdOut in listOut)
            {
                foreach(JMD jmd in jmds)
                {
                    if (jmd.PL.ObjectId == jmdOut.PL.ObjectId)
                    {
                        return true;
                    }
                }
              
            }
            return false;

        }

        private void SetQTGS(JMD jmd, string zy)
        {
            jmd.QTGSB = zy;
            jmd.QTGSD = zy;
            jmd.QTGSX = zy;
            jmd.QTGSN = zy;
        }

        private  IList<Polyline> SelectContainsFWX1(Zdinfo zdinfo)
        {
            IList<Polyline> list = new List<Polyline>();
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(zdinfo.PC);
            Autodesk.AutoCAD.DatabaseServices.Region region = CADUtils.CreateRegion(zdinfo.Polyline);
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline = zdinfo.Polyline;
                    Polyline polyline2 = (Polyline)dBObject;
                    if (!(polyline == polyline2))
                    {
                        Autodesk.AutoCAD.DatabaseServices.Region region2 = CADUtils.CreateRegion(polyline2);
                        if (region2 != null)
                        {
                            double area = region.Area;
                            region.BooleanOperation(0, region2);
                            if (region.Area < area + 0.1)
                            {
                                JMD jMD = PLToJMD(polyline2, FWDCB_FW);
                                if (jMD != null)
                                {
                                    list.Add(polyline2);
                                }
                            }
                        }
                    }
                }
            }
            return list;
        }

        private static IList<JMD> PLToJMD(DBObjectCollection dbc, Dictionary<string, string> dic)
        {
            IList<JMD> list = new List<JMD>();
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is Polyline)
                {
                    Polyline jmdPL = (Polyline)dBObject;
                    JMD jMD = PLToJMD(jmdPL, dic);
                    if (jMD != null)
                    {
                        list.Add(jMD);
                    }
                }
            }
            return list;
        }

        private static JMD PLToJMD(Polyline jmdPL, Dictionary<string, string> FWDCB_FW)
        {
            ResultBuffer xDataForApplication = jmdPL.GetXDataForApplication("SOUTH");
            JMD result;
            if (xDataForApplication != null)
            {
                TypedValue[] array = xDataForApplication.AsArray();
                if (array.Length <= 3)
                {
                    string name;
                    if (FWDCB_FW.TryGetValue(array[1].Value.ToString(), out name))
                    {
                        result = new JMD();
                        result.Name = name;
                        result.PL = jmdPL;

                        return result;
                    }
                }
            }
            result = null;
            return result;
        }

        private  IWorkbook ExportFWDCB(Zdinfo zdinfo, IList<JMD> jdms)
        {
            string zdnum = zdinfo.getZdnum();
            ZDX zdx = zdinfo.zdx;
            IWorkbook result;
            if (zdx == null)
            {
                result = null;
            }
            else
            {
                if (zdnum == null || zdnum.Length != 19)
                {
                    MessageBox.Show("宗地编码不是19位：" + zdnum);
                }
                IWorkbook workbook = ExcelRead.ReadExcel(FWDCB);
                ISheet sheetAt = workbook.GetSheetAt(0);
                IRow row = sheetAt.GetRow(5);
                ExcelWrite.SetValue(row, 0, zdnum.Substring(0, 6));
                ExcelWrite.SetValue(row, 4, zdnum.Substring(6, 3));
                ExcelWrite.SetValue(row, 7, zdnum.Substring(9, 3));
                ExcelWrite.SetValue(row, 10, zdnum.Substring(12, 7));
                ExcelWrite.SetValue(row, 13, zdx.DZWFWDM);


                ExcelWrite.SetValue(sheetAt.GetRow(7), 4, zdx.BDCDYH);
                ExcelWrite.SetValue(sheetAt.GetRow(10), 4, zdx.TXDZ);
                ExcelWrite.SetValue(sheetAt.GetRow(10), 11, zdx.YZBM);
                ExcelWrite.SetValue(sheetAt.GetRow(13), 4, zdinfo.getQuanli());
                ExcelWrite.SetValue(sheetAt.GetRow(13), 13, zdx.ZJZL);
                ExcelWrite.SetValue(sheetAt.GetRow(15), 13, zdx.SFZHM);
                ExcelWrite.SetValue(sheetAt.GetRow(17), 1, zdx.LXDH);
                ExcelWrite.SetValue(sheetAt.GetRow(17), 7, zdx.TXDZ);
                ExcelWrite.SetValue(sheetAt.GetRow(20), 1, zdx.QLRLX);
                ExcelWrite.SetValue(sheetAt.GetRow(22), 1, zdx.FWXZ);
                ExcelWrite.SetValue(sheetAt.GetRow(24), 1, zdx.YT);
                ExcelWrite.SetValue(sheetAt.GetRow(24), 7, zdx.GHYT);
                string value = DateTime.Now.ToString("yyyy年M月dd日");
                //ExcelWrite.SetValue(sheetAt.GetRow(41), 11, value);
                int count = jdms.Count;
                if (count > 3)
                {
                    ExcelWrite.CreateRows(sheetAt, 32, 17, 32, count - 3);
                    sheetAt.AddMergedRegion(new CellRangeAddress(27, 34 + count - 3, 0, 0));
                }
                JMDWriteCellValue(zdx, jdms, sheetAt, 32);
                result = workbook;
            }
            return result;
        }

        private static void JMDWriteCellValue(ZDX zdx, IList<JMD> jmds, ISheet sheet, int rowStart)
        {

            for (int i = 0; i < jmds.Count; i++)
            {
                JMD jMD = jmds[i];
                IRow row = sheet.GetRow(i + rowStart);
                for (int j = 0; j < 17; j++)
                {
                    ICell cell = row.GetCell(j);
                    switch (j)
                    {
                        case 1:

                            cell.SetCellValue("F000"+(i+1));
                            break;
                        case 2:
                            cell.SetCellValue("F0001");
                            break;
                        case 3:
                            cell.SetCellValue(1.0);
                            break;
                        case 4:
                            cell.SetCellValue(1.0);
                            break;
                        case 5:
                            cell.SetCellValue(1.0);
                            break;
                        case 6:
                            cell.SetCellValue(jMD.Name);
                            break;
                        case 8:
                        case 9:
                        case 10:
                            cell.SetCellValue(jMD.PL.Area.ToString("f2"));
                            break;
                        case 7:
                        case 11:
                        case 12:
                            cell.SetCellValue("/");
                            break;
                        case 13:
                            cell.SetCellValue(jMD.QTGSD);
                            
                            break;
                        case 14:
                            cell.SetCellValue(jMD.QTGSN);
                            
                            break;
                        case 15:
                            cell.SetCellValue(jMD.QTGSX);
                            
                            break;
                        case 16:
                            cell.SetCellValue(jMD.QTGSB);
                            break;
                    }
                }
            }
        }

        public static IList<Zdinfo> DBCToZdinfo(DBObjectCollection dbc)
        {
            IList<Zdinfo> list = new List<Zdinfo>();
            foreach (DBObject dbo in dbc)
            {
                Zdinfo zdinfo = ToZdinfo(dbo);
                if (zdinfo != null)
                {
                    list.Add(zdinfo);
                }
            }
            return list;
        }

        internal void ExportQJDCB(string saveDir)
        {
            DBObjectCollection dbc = SelectJZD("选择要导出权籍调查表的权属线\n");
            IList<Zdinfo> list = DBCToZdinfo(dbc);
            string mdbPath = FileUtils.SelectSingleFile("打开mdb", "Access数据库(*.mdb)|*.mdb");
            Dictionary<string, List<Jzxinfo>> dic = OperateMdb.MdbToJzxinfo(mdbPath);
            this.ZdinfoSetJzx(list, dic);
            this.SetJzxNull(list);
            foreach (Zdinfo current in list)
            {
                IWorkbook workbook = ExcelRead.ReadExcel(QJDCB);
                ISheet sheet = workbook.GetSheet("界址说明表");
                IRow row = sheet.GetRow(1);
                ICell cell = row.GetCell(2);
                string cellValue = this.PointDiscrtion(current);
                cell.SetCellValue(cellValue);
                string cellValue2 = this.PolyDisction(current);
                row = sheet.GetRow(19);
                cell = row.GetCell(2);
                cell.SetCellValue(cellValue2);
                sheet = workbook.GetSheet("界址签章表");
                IList<string[]> qZBContent = this.GetQZBContent(current);
                this.QZBWrite(sheet, qZBContent);
                ExcelWrite.Save(workbook, string.Concat(new string[]
				{
					saveDir,
					"\\",
					current.getZdnum(),
					"_",
					current.getQuanli(),
					"_不动产权籍调查表.xls"
				}));
            }
        }

        private void QZBWrite(ISheet sheet, IList<string[]> qzbContect)
        {
            int num = 7;
            foreach (string[] current in qzbContect)
            {
                IRow row = sheet.GetRow(num++);
                for (int i = 0; i < 4; i++)
                {
                    ICell cell = row.GetCell(i);
                    string text = current[i];
                    cell.SetCellValue(text);
                    if (text != null && text.Length > 39)
                    {
                        row.HeightInPoints = (float)(text.Length + 20);
                    }
                }
            }
        }

        public void SetJzxNull(IList<Zdinfo> zds)
        {
            foreach (Zdinfo current in zds)
            {
                this.SetJzxNull(current);
            }
        }

        public void SetJzxNull(Zdinfo zd)
        {
            IList<Jzxinfo> list = zd.JZXS;
            int num = zd.PC.Count - 1;
            if (list == null)
            {
                list = new List<Jzxinfo>();
                for (int i = 0; i <= num; i++)
                {
                    Jzxinfo jzxinfo = new Jzxinfo();
                    if (i == num)
                    {
                        jzxinfo.setQdh("J1");
                    }
                    else
                    {
                        jzxinfo.setQdh("J" + (i + 1));
                    }
                }
            }
        }

        private void ZdinfoSetJzx(IList<Zdinfo> zds, Dictionary<string, List<Jzxinfo>> dic)
        {
            foreach (Zdinfo current in zds)
            {
                string zdnum = current.getZdnum();
                List<Jzxinfo> list;
                if (dic.TryGetValue(zdnum, out list))
                {
                    if (list.Count < current.PC.Count)
                    {
                        throw new Exception("宗地编码是：" + list[0].getBzdh() + "权属线点有重复，请先删除重复点");
                    }
                    Jzxinfo jzxinfo = new Jzxinfo();
                    jzxinfo.setQdh(list[0].getQdh());
                    list.Add(jzxinfo);
                    current.JZXS = list;
                }
            }
        }
        
        public void fun(Dictionary<bool, int[]> dic, Polyline polyline)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            polyline.GetStretchPoints(point3dCollection);
            

        }
        private IList<string[]> GetQZBContent(Zdinfo zd)
        {
            Polyline polyline = zd.Polyline;
            Point3dCollection point3dCollection = new Point3dCollection();
            polyline.GetStretchPoints(point3dCollection);
            double[] dou = utils.FindFourPoints(point3dCollection);
            Point3dCollection point3dCollection2 = utils.createFourPoint(dou);
            Polyline3d polyline3d = utils.CreatePolye(point3dCollection2, true);
            IList<Zdinfo> list = DBCToZdinfo(utils.SelectJzd(point3dCollection2));
            int num = utils.PolyDiPoly(point3dCollection, point3dCollection2);
            int[] array = new int[4];
            if (num > 3)
            {
                int[] array2 = utils.FindPointsDic(point3dCollection);
                if (array2[0] == 0)
                {
                    array = array2;
                }
                else
                {
                    int[] array3 = utils.FindPointsDic(point3dCollection, point3dCollection2);
                    if (array3[0] == 0)
                    {
                        array = array3;
                    }
                    else
                    {
                        int[][] max = new int[][]
						{
							array2,
							array3
						};
                        int num2 = utils.SelectPc1PC2(max, point3dCollection);
                        if (num2 == 0)
                        {
                            array = array2;
                        }
                        else if (num2 == 1)
                        {
                            array = array3;
                        }
                    }
                }
            }
            else
            {
                array = utils.FindPointsDic(point3dCollection, point3dCollection2);
            }
            IList<Jzxinfo> jZXS = zd.JZXS;
            IList<string[]> list2 = new List<string[]>();
            double num3 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[1]]);
            double num4 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[3]]);
            double num5 = num3 - num4;
            for (int i = 0; i < 4; i++)
            {
                int start = array[i];
                int end = array[i + 1];
                DBObjectCollection dbc = this.SelectJzd(polyline, point3dCollection, array, i, 0.02, 0.99, 1);
                IList<Zdinfo> list3 = DBCToZdinfo(dbc);
                string[] qZB_DH = this.GetQZB_DH(start, end, jZXS);
                string text = "";
                if (num5 < 0.0 && num3 < 40.0)
                {
                    switch (i)
                    {
                        case 0:
                            text = zd.getBz();
                            break;
                        case 1:
                            text = zd.getDz();
                            break;
                        case 2:
                            text = zd.getNz();
                            break;
                        case 3:
                            text = zd.getXz();
                            break;
                    }
                }
                else
                {
                    switch (i)
                    {
                        case 0:
                            text = zd.getDz();
                            break;
                        case 1:
                            text = zd.getNz();
                            break;
                        case 2:
                            text = zd.getXz();
                            break;
                        case 3:
                            text = zd.getBz();
                            break;
                    }
                }
                if (list3.Count > 1)
                {
                    string zdnum = zd.getZdnum();
                    foreach (Zdinfo current in list3)
                    {
                        string zdnum2 = current.getZdnum();
                        if (!zdnum2.Equals(zdnum))
                        {
                            if (text != null && text.Equals(current.getQuanli()))
                            {
                                if (zdnum2.Length > 18)
                                {
                                    text = text + "\r\n" + zdnum2.Substring(12, 7);
                                }
                            }
                        }
                    }
                }
                qZB_DH[3] = text;
                list2.Add(qZB_DH);
            }
            return list2;
        }

        private string[] GetQZB_DH(int start, int end, IList<Jzxinfo> jzxinfo)
        {
            string[] array = new string[4];
            array[0] = jzxinfo[start].getQdh();
            array[2] = jzxinfo[end].getQdh();
            StringBuilder stringBuilder = new StringBuilder();
            if (start < end)
            {
                start++;
                while (start < end)
                {
                    stringBuilder.Append(jzxinfo[start++].getQdh() + "、");
                }
            }
            else
            {
                int num = jzxinfo.Count - 2;
                while (start < num)
                {
                    stringBuilder.Append(jzxinfo[++start].getQdh() + "、");
                }
                start = 0;
                while (start < end)
                {
                    stringBuilder.Append(jzxinfo[start++].getQdh() + "、");
                }
            }
            if (stringBuilder.Length > 1)
            {
                stringBuilder.Remove(stringBuilder.Length - 1, 1);
            }
            array[1] = stringBuilder.ToString();
            return array;
        }

        private string PolyDisction(Zdinfo zd)
        {
            Point3dCollection pC = zd.PC;
            StringBuilder stringBuilder = new StringBuilder();
            int num = pC.Count - 1;
            IList<Jzxinfo> jZXS = zd.JZXS;
            for (int i = 0; i <= num; i++)
            {
                Jzxinfo jzxinfo = jZXS[i];
                Point3d point3d = pC[i];
                stringBuilder.Append(jzxinfo.getQdh() + "-");
                Point3d point3d2;
                if (i < num)
                {
                    point3d2 = pC[i + 1];
                    stringBuilder.Append(jZXS[i + 1].getQdh());
                }
                else
                {
                    point3d2 = pC[0];
                    stringBuilder.Append(jZXS[i + 1].getQdh());
                }
                stringBuilder.Append("，" + jzxinfo.getQdh() + "沿");
                Point3dCollection containsPL = this.GetContainsPL(point3d, point3d2);
                JMD resultJMD = this.GetResultJMD(zd, containsPL);
                string str = "外";
                string str2;
                if (resultJMD == null)
                {
                    str2 = "两点连线";
                }
                else
                {
                    str2 = resultJMD.Name;
                    str = resultJMD.Loaction;
                }
                stringBuilder.Append(str2 + str);
                string str3 = this.XiangXian(this.PolylinXiangXian(point3d, point3d2));
                stringBuilder.Append(str3 + "方向" + this.PointLegth(point3d, point3d2).ToString("f2") + "米");
                if (i < num)
                {
                    stringBuilder.Append(jZXS[i + 1].getQdh() + ";\r\n");
                }
                else
                {
                    stringBuilder.Append(jZXS[i + 1].getQdh() + "。");
                }
            }
            return stringBuilder.ToString();
        }

        private Point3dCollection GetContainsPL(Point3d pt1, Point3d pt2)
        {
            double num = this.PointLegth(pt1, pt2);
            if (num > 0.31)
            {
                num = 0.3;
            }
            else
            {
                num = 0.1;
            }
            Point3dCollection point3dCollection = new Point3dCollection();
            double x = pt1.X;
            double y = pt1.Y;
            double x2 = pt2.X;
            double y2 = pt2.Y;
            point3dCollection.Add(new Point3d(x + num, y + num, 0.0));
            point3dCollection.Add(new Point3d(x2 - num, y2 + num, 0.0));
            point3dCollection.Add(new Point3d(x2 - num, y2 - num, 0.0));
            point3dCollection.Add(new Point3d(x + num, y - num, 0.0));
            return CADUtils.PointOffset(pt1, pt2, 0.05);
        }

        private string PointDiscrtion(Zdinfo zd)
        {
            Point3dCollection pC = zd.PC;
            StringBuilder stringBuilder = new StringBuilder();
            IList<Jzxinfo> jZXS = zd.JZXS;
            for (int i = 0; i < pC.Count; i++)
            {
                Jzxinfo jzxinfo = jZXS[i];
                IList<JMD> list = this.SelectFWX(pC[i]);
                JMD resultJMD = this.GetResultJMD(zd, pC[i]);
                string str;
                if (resultJMD == null)
                {
                    str = "位于拐点";
                }
                else
                {
                    str = this.PointDiscrtion(pC[i], resultJMD);
                }
                stringBuilder.Append(jzxinfo.getQdh() + str + "；");
            }
            stringBuilder.Remove(stringBuilder.Length - 1, 1);
            stringBuilder.Append("。");
            return stringBuilder.ToString();
        }

        public JMD GetResultJMD(Zdinfo zd, Point3d point3d)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            double x = point3d.X;
            double y = point3d.Y;
            point3dCollection.Add(new Point3d(x - 0.1, y + 0.1, 0.0));
            point3dCollection.Add(new Point3d(x + 0.1, y + 0.1, 0.0));
            point3dCollection.Add(new Point3d(x, y - 0.1, 0.0));
            return this.GetResultJMD(zd, point3dCollection);
        }

        public JMD GetResultJMD(Zdinfo zd, Point3dCollection pc)
        {
            IList<JMD> list = this.SelectFWX(pc);
            JMD result;
            if (list == null || list.Count == 0)
            {
                result = null;
            }
            else
            {
                IList<JMD> list2 = this.SelectContanisDWS(list, zd.Polyline);
                JMD jMD;
                if (list2 == null || list2.Count == 0)
                {
                    jMD = this.SelectFWX(list, "房屋");
                    if (jMD == null)
                    {
                        jMD = list[0];
                    }
                    jMD.Loaction = "内";
                }
                else
                {
                    jMD = this.SelectFWX(list2, "墙壁");
                    if (jMD == null)
                    {
                        jMD = list2[0];
                    }
                    jMD.Loaction = "外";
                }
                result = jMD;
            }
            return result;
        }

        private JMD SelectFWX(IList<JMD> jmds, string dw)
        {
            JMD result;
            foreach (JMD current in jmds)
            {
                if (current.Name.Equals(dw))
                {
                    result = current;
                    return result;
                }
            }
            result = null;
            return result;
        }

        private string PointDiscrtion(Point3d point3d, JMD jmd)
        {
            string name = jmd.Name;
            string result;
            if (name.Equals("墙壁"))
            {
                string str = this.PLPointFangXinag(point3d, jmd.PL);
                result = "位于房屋" + str + "角";
            }
            else
            {
                result = "位于" + name + "拐点";
            }
            return result;
        }

        private string PLPointFangXinag(Point3d point, Polyline polyline)
        {
            Point3d polyCore = utils.GetPolyCore(polyline);
            int xinagXian = this.PolylinXiangXian(polyCore, point);
            return this.XiangXian(xinagXian);
        }

        private string XiangXian(int xinagXian)
        {
            string result = "";
            switch (xinagXian)
            {
                case 1:
                    result = "西北";
                    break;
                case 2:
                    result = "西南";
                    break;
                case 3:
                    result = "东南";
                    break;
                case 4:
                    result = "东北";
                    break;
            }
            return result;
        }

        public double ComputeAngle(Point3d nowpoint, Point3d CentPoint)
        {
            double num = this.PointLegth(nowpoint, CentPoint);
            double num2 = Math.Asin(Math.Abs(nowpoint.Y - CentPoint.Y) / num);
            double num3 = num2 * 180.0 / 3.1415926535897931;
            if (CentPoint.X - nowpoint.X <= 0.0 && CentPoint.Y - nowpoint.Y >= 0.0)
            {
                num3 = 90.0 - num3;
            }
            else if (CentPoint.X - nowpoint.X <= 0.0 && CentPoint.Y - nowpoint.Y <= 0.0)
            {
                num3 += 90.0;
            }
            else if (CentPoint.X - nowpoint.X >= 0.0 && CentPoint.Y - nowpoint.Y <= 0.0)
            {
                num3 = 270.0 - num3;
            }
            else if (CentPoint.X - nowpoint.X >= 0.0 && CentPoint.Y - nowpoint.Y >= 0.0)
            {
                num3 += 270.0;
            }
            return num3 - 235.0;
        }

        public int PolylinXiangXian(Point3d nowpoint, Point3d CentPoint)
        {
            double num = this.PointLegth(nowpoint, CentPoint);
            double num2 = Math.Asin(Math.Abs(nowpoint.Y - CentPoint.Y) / num);
            double num3 = num2 * 180.0 / 3.1415926535897931;
            int result;
            if (CentPoint.X - nowpoint.X <= 0.0 && CentPoint.Y - nowpoint.Y >= 0.0)
            {
                result = 1;
            }
            else if (CentPoint.X - nowpoint.X <= 0.0 && CentPoint.Y - nowpoint.Y <= 0.0)
            {
                result = 2;
            }
            else if (CentPoint.X - nowpoint.X >= 0.0 && CentPoint.Y - nowpoint.Y <= 0.0)
            {
                result = 3;
            }
            else if (CentPoint.X - nowpoint.X >= 0.0 && CentPoint.Y - nowpoint.Y >= 0.0)
            {
                result = 4;
            }
            else
            {
                result = 0;
            }
            return result;
        }

        public double PointLegth(Point3d pa, Point3d pb)
        {
            return Math.Sqrt(Math.Pow(pa.X - pb.X, 2.0) + Math.Pow(pa.Y - pb.Y, 2.0));
        }

        private IList<JMD> SelectContanisDWS(IList<JMD> dws, Polyline polyline)
        {
            IList<JMD> result;
            if (dws == null || dws.Count == 0)
            {
                result = null;
            }
            else
            {
                IList<JMD> list = new List<JMD>();
                foreach (JMD current in dws)
                {
                    if (CADUtils.CheckPLContains(polyline, current.PL))
                    {
                        list.Add(current);
                    }
                }
                result = list;
            }
            return result;
        }

        private IList<JMD> SelectFWX(Point3d point3d)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            double x = point3d.X;
            double y = point3d.Y;
            point3dCollection.Add(new Point3d(x - 0.1, y + 0.1, 0.0));
            point3dCollection.Add(new Point3d(x + 0.1, y + 0.1, 0.0));
            point3dCollection.Add(new Point3d(x, y - 0.1, 0.0));
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(point3dCollection, types);
            return PLToJMD(corssEntityCollection, FWDCB_PL);
        }

        private IList<JMD> SelectFWX(Point3dCollection pc)
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(pc, types);
            return PLToJMD(corssEntityCollection, FWDCB_PL);
        }

        internal IWorkbook ExportJzdsExcel(DBObjectCollection dbc)
        {
            IList<Zdinfo> list = DBCToZdinfo(dbc);
            IWorkbook result;
            if (list.Count == 0)
            {
                result = null;
            }
            else
            {
                result = this.ExportJzdsExcel(list);
            }
            return result;
        }

        internal IWorkbook ExportJzdsExcel(IList<Zdinfo> zds)
        {
            IWorkbook workbook = ExcelRead.ReadExcel("C:\\ZJDTemplete\\宅基地文件\\统计界址点个数模板.xls");
            ISheet sheetAt = workbook.GetSheetAt(0);
            for (int i = 0; i < zds.Count; i++)
            {
                Zdinfo zdinfo = zds[i];
                IRow row = sheetAt.CreateRow(i + 1);
                for (int j = 0; j < 3; j++)
                {
                    ICell cell = row.CreateCell(j);
                    switch (j)
                    {
                        case 0:
                            cell.SetCellValue(zdinfo.getZdnum());
                            break;
                        case 1:
                            cell.SetCellValue(zdinfo.getQuanli());
                            break;
                        case 2:
                            cell.SetCellValue((double)zdinfo.PC.Count);
                            break;
                    }
                }
            }
            return workbook;
        }

        internal IWorkbook ZDTExportJzdsExcel()
        {
            IList<Zdinfo> list = this.findZDT_Zdinfo();
            IWorkbook result;
            if (list.Count == 0)
            {
                result = null;
            }
            else
            {
                result = this.ExportJzdsExcel(list);
            }
            return result;
        }

        private IList<Zdinfo> findZDT_Zdinfo()
        {
            PlanMapService planMapService = new PlanMapService();
            Dictionary<string, DBObjectCollection> zdDic_zdnum = planMapService.GetZdDic_zdnum19();
            IList<Zdinfo> list = new List<Zdinfo>();
            foreach (string current in zdDic_zdnum.Keys)
            {
                IList<Zdinfo> list2 = DBCToZdinfo(zdDic_zdnum[current]);
                foreach (Zdinfo current2 in list2)
                {
                    if (current.Equals(current2.getZdnum()))
                    {
                        list.Add(current2);
                        break;
                    }
                }
            }
            return list;
        }
    }
}
