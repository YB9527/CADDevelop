using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


namespace YanBo
{
    internal class Utils
    {
        private static Dictionary<string, string> dic;

        private static Dictionary<string, string> polyDic;

        private static Dictionary<string, string> polyFwDic;

        private static Dictionary<string, string> txDic = Utils.GetTxSurface();

        public static string CADDir()
        {
            return new Utils().GetDocument().Name;
        }

        public static void CheckXZDMTable(Dictionary<string, string> dic)
        {
            Dictionary<string, string>.KeyCollection keys = dic.Keys;
            IList<string> localXzdm = Utils.GetLocalXzdm();
            foreach (string current in keys)
            {
                if (current.Length > 8)
                {
                    if (!localXzdm.Contains(current.Substring(0, 9)))
                    {
                        if (!new Main().getUseDate().Contains("永久"))
                        {
                            throw new Exception("你的行政代表不是本项目的：" + current);
                        }
                    }
                }
            }
        }

        public static IList<String> GetLocalXzdm()
        {
            IList<String> list = new List<String>();
            list.Add("510185039");
            list.Add("510185047");
            list.Add("510185037");
            list.Add("510185042");
            list.Add("510185045");
            list.Add("DJZQDM");

            return list;
        }
      
        public void ReadData()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Entity entity = null;
            PromptEntityResult ent = ed.GetEntity("\n选择要读取数据的对象");
            if (ent.Status == PromptStatus.OK)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    entity = (Entity)trans.GetObject(ent.ObjectId, OpenMode.ForRead, true);
                    DBDictionary extensionDic = trans.GetObject(entity.ExtensionDictionary, OpenMode.ForRead) as DBDictionary;
                    DataTable dt = trans.GetObject(extensionDic.GetAt(""), OpenMode.ForWrite) as DataTable;

                }

            }
        }

        public int[] FindPointsDic(Point3dCollection points, Point3dCollection fourPoints)
        {
            int[] array = new int[5];
            double num = 100000.0;
            for (int i = 0; i < points.Count; i++)
            {
                if (num > this.CalculateTwoPt(points[i], fourPoints[0]))
                {
                    num = this.CalculateTwoPt(points[i], fourPoints[0]);
                    array[0] = i;
                    array[4] = i;
                }
            }
            double num2 = 100000.0;
            for (int i = 0; i < points.Count; i++)
            {
                if (num2 > this.CalculateTwoPt(points[i], fourPoints[1]))
                {
                    num2 = this.CalculateTwoPt(points[i], fourPoints[1]);
                    array[1] = i;
                }
            }
            double num3 = 100000.0;
            for (int i = 0; i < points.Count; i++)
            {
                if (num3 > this.CalculateTwoPt(points[i], fourPoints[2]))
                {
                    num3 = this.CalculateTwoPt(points[i], fourPoints[2]);
                    array[2] = i;
                }
            }
            double num4 = 100000.0;
            for (int i = 0; i < points.Count; i++)
            {
                if (num4 > this.CalculateTwoPt(points[i], fourPoints[3]))
                {
                    num4 = this.CalculateTwoPt(points[i], fourPoints[3]);
                    array[3] = i;
                }
            }
            return array;
        }

        public int[] FindPointsDic(Point3dCollection points)
        {
            int[] array = new int[5];
            double[] array2 = new double[5];
            array2[0] = points[0].X;
            array2[1] = points[0].X;
            array2[2] = points[0].Y;
            array2[3] = points[0].Y;
            for (int i = 1; i < points.Count; i++)
            {
                if (array2[0] >= points[i].X)
                {
                    array2[0] = points[i].X;
                    array[3] = i;
                }
            }
            for (int i = 1; i < points.Count; i++)
            {
                if (array2[1] < points[i].X)
                {
                    array2[1] = points[i].X;
                    array[1] = i;
                }
            }
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i].Y > array2[2])
                {
                    array2[2] = points[i].Y;
                    array[0] = i;
                    array[4] = i;
                }
            }
            for (int i = 0; i < points.Count; i++)
            {
                if (array2[3] > points[i].Y)
                {
                    array2[3] = points[i].Y;
                    array[2] = i;
                }
            }
            return array;
        }

        public double[] FindFourPoints(Point3dCollection points)
        {
            double[] array = new double[]
			{
				points[0].X,
				points[0].X,
				points[0].Y,
				points[0].Y
			};
            for (int i = 1; i < points.Count; i++)
            {
                if (array[0] > points[i].X)
                {
                    array[0] = points[i].X;
                }
            }
            for (int i = 1; i < points.Count; i++)
            {
                if (array[1] < points[i].X)
                {
                    array[1] = points[i].X;
                }
            }
            for (int i = 0; i < points.Count; i++)
            {
                if (points[i].Y > array[2])
                {
                    array[2] = points[i].Y;
                }
            }
            for (int i = 0; i < points.Count; i++)
            {
                if (array[3] > points[i].Y)
                {
                    array[3] = points[i].Y;
                }
            }
            return array;
        }

        public Point3dCollection createFourPoint(double[] dou)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            Point3d point3d = new Point3d(dou[0], dou[2], 0.0);
            Point3d point3d2 = new Point3d(dou[1], dou[2], 0.0);
            Point3d point3d3 = new Point3d(dou[1], dou[3], 0.0);
            Point3d point3d4 = new Point3d(dou[0], dou[3], 0.0);
            point3dCollection.Add(point3d);
            point3dCollection.Add(point3d2);
            point3dCollection.Add(point3d3);
            point3dCollection.Add(point3d4);
            return point3dCollection;
        }

        public Point3dCollection CreatePolye(Point3dCollection points)
        {
            Polyline3d polyline3d = this.CreatePolye(points, true);
            Point3dCollection point3dCollection = new Point3dCollection();
            polyline3d.GetStretchPoints(point3dCollection);
            return point3dCollection;
        }

        public Polyline3d CreatePolye(Point3dCollection points, bool flag)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            foreach (Point3d point3d in points)
            {
                point3dCollection.Add(new Point3d(point3d.X, point3d.Y, 0.0));
            }
            points = point3dCollection;
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Polyline3d polyline3d = new Polyline3d(0, points, true);
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                polyline3d.ColorIndex=(2);
                Utils.CreateLayer("删除");
                polyline3d.Layer=("删除");
                blockTableRecord.AppendEntity(polyline3d);
                transaction.AddNewlyCreatedDBObject(polyline3d, true);
                transaction.Commit();
            }
            return polyline3d;
        }

        public Point3dCollection findEntityPoints(string select)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            DBObjectCollection dBObjectCollection = this.findTexts(select);
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in dBObjectCollection)
                {
                    if (dBObject is DBText)
                    {
                        DBText dBText = (DBText)transaction.GetObject(dBObject.ObjectId, OpenMode.ForWrite, true);
                        point3dCollection.Add(dBText.Position);
                    }
                }
                transaction.Commit();
            }
            return point3dCollection;
        }

        public DBObjectCollection findTexts(string select)
        {
            DBObjectCollection dBObjectCollection = new DBObjectCollection();
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "TEXT")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionResult selection = editor.GetSelection(selectionFilter);
            DBObjectCollection entityCollection = this.GetEntityCollection(selection);
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in entityCollection)
                {
                    DBText dBText = (DBText)transaction.GetObject(dBObject.ObjectId, OpenMode.ForWrite, true);
                    if (dBText.TextString.Equals(select))
                    {
                        dBObjectCollection.Add(dBText);
                    }
                }
                transaction.Commit();
            }
            return dBObjectCollection;
        }

        public DBObjectCollection GetCorssEntityCollection(Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            PromptSelectionResult ents = this.GetEditor().SelectCrossingPolygon(point3dCollection);
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection SelectJzd(Point3dCollection pc)
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            return this.GetCorssEntityCollection(pc, types);
        }

        public DBObjectCollection GetCorssEntityCollection(Polyline3d poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            PromptSelectionResult ents = this.GetEditor().SelectCrossingPolygon(point3dCollection);
            
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection GetCorssEntityCollection(Point3dCollection pc)
        {
            PromptSelectionResult ents = GetEditor().SelectCrossingPolygon(pc);
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection GetCorssEntityCollection(Polyline3d poly, TypedValue[] types)
        {
            SelectionFilter selectionFilter = new SelectionFilter(types);
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            PromptSelectionResult ents = this.GetEditor().SelectCrossingPolygon(point3dCollection, selectionFilter);
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection GetCorssEntityCollection(Polyline poly, TypedValue[] types)
        {
            SelectionFilter selectionFilter = new SelectionFilter(types);
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            PromptSelectionResult ents = this.GetEditor().SelectCrossingPolygon(point3dCollection, selectionFilter);
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection GetCorssEntityCollection(Point3dCollection pc, TypedValue[] types)
        {
            SelectionFilter selectionFilter = new SelectionFilter(types);
            PromptSelectionResult ents = this.GetEditor().SelectCrossingPolygon(pc, selectionFilter);
            return this.GetEntityCollection(ents);
        }

        public DBObjectCollection GetEntityCollection(PromptSelectionResult ents)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            DBObjectCollection dBObjectCollection = new DBObjectCollection();
            if (ents.Status == PromptStatus.OK)
            {
                using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                {
                    SelectionSet value = ents.Value;
                    ObjectId[] objectIds = value.GetObjectIds();
                    for (int i = 0; i < objectIds.Length; i++)
                    {
                        ObjectId objectId = objectIds[i];
                        Entity entity = transaction.GetObject(objectId, OpenMode.ForWrite, true) as Entity;
                        if (entity != null)
                        {
                            dBObjectCollection.Add(entity);
                        }
                    }
                    transaction.Commit();
                }
            }
            return dBObjectCollection;
        }

        public DBText getTextEntity(string tips)
        {
            PromptEntityResult entity = Application.DocumentManager.MdiActiveDocument.Editor.GetEntity(tips);
            DBText result;
            if (entity.Status == PromptStatus.OK)
            {
                using (Transaction transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
                {
                    Entity entity2 = transaction.GetObject(entity.ObjectId, OpenMode.ForWrite, true) as Entity;
                    if (entity2 is DBText)
                    {
                        DBText dBText = entity2 as DBText;
                        result = dBText;
                        return result;
                    }
                    transaction.Commit();
                }
            }
            result = null;
            return result;
        }

        public Point3d getLineXmin(Polyline3d poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            Point3d result = point3dCollection[0];
            foreach (Point3d point3d in point3dCollection)
            {
                if (result.X > point3d.X)
                {
                    result = point3d;
                }
            }
            return result;
        }

        public Document GetDocument()
        {
            return Application.DocumentManager.MdiActiveDocument;
        }

        public Database GetDabase()
        {
            return HostApplicationServices.WorkingDatabase;
        }

        public Editor GetEditor()
        {
            return this.GetDocument().Editor;
        }

        public static void CreateLayer(string layer)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Transaction transaction = workingDatabase.TransactionManager.StartTransaction();
            try
            {
                LayerTable layerTable = (LayerTable)transaction.GetObject(workingDatabase.LayerTableId, OpenMode.ForWrite, true);
                if (layerTable.Has(layer))
                {
                    ObjectId objectId = layerTable[layer];
                }
                else
                {
                    LayerTableRecord layerTableRecord = new LayerTableRecord();
                    layerTableRecord.Name=(layer);
                    ObjectId objectId = layerTable.Add(layerTableRecord);
                    transaction.AddNewlyCreatedDBObject(layerTableRecord, true);
                }
                transaction.Commit();
            }
            catch
            {
                transaction.Abort();
            }
            finally
            {
                transaction.Dispose();
            }
        }

        public static Dictionary<string, string> GetSurface()
        {
            if (Utils.dic == null)
            {
                Utils.dic = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\编码块地物.xls", 0);
            }
            return Utils.dic;
        }

        public static Dictionary<string, string> GetPolySurface()
        {
            if (Utils.polyDic == null)
            {
                Utils.polyDic = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\编码线地物.xls", 0);
            }
            return Utils.polyDic;
        }

        public static Dictionary<string, string> GetFw_PolySurface()
        {
            if (Utils.polyFwDic == null)
            {
                Utils.polyFwDic = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\编码界址线类别.xls", 0);
            }
            return Utils.polyFwDic;
        }

        public static Dictionary<string, string> GetTxSurface()
        {
            if (Utils.txDic == null)
            {
                Utils.txDic = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\编码文字地物.xls", 0);
            }
            return Utils.txDic;
        }

        public object[] GetSurfacePoly3d(Polyline3d poly, Polyline basicPoly)
        {
            DBObjectCollection corssEntityCollection = this.GetCorssEntityCollection(poly);
            object[] result;
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline poly2 = (Polyline)dBObject;
                    object[] surface = this.GetSurface(poly2, basicPoly);
                    if (surface != null)
                    {
                        result = surface;
                        return result;
                    }
                }
            }
            result = null;
            return result;
        }

        public TypedValue[] GetQSSurface(DBObjectCollection dbc)
        {
            TypedValue[] array = null;
            TypedValue[] result;
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is Polyline)
                {
                    TypedValue[] surQSface = this.GetSurQSface((Polyline)dBObject);
                    if (surQSface != null)
                    {
                        if (array != null)
                        {
                            System.Windows.Forms.MessageBox.Show("宗地号是：" + array[2].Value.ToString() + ",处有两个未封闭的权属线");
                            result = array;
                            return result;
                        }
                        array = (TypedValue[])surQSface.Clone();
                    }
                }
            }
            result = array;
            return result;
        }

        public TypedValue[] GetSurQSface(Polyline poly)
        {
            Dictionary<string, string> surface = Utils.GetSurface();
            ResultBuffer resultBuffer = poly.XData;
            TypedValue[] result;
            if (resultBuffer == null || !poly.Closed)
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

        public object[] GetSurface(Polyline poly, Polyline basicPoly)
        {
            object[] result;
            if (poly == basicPoly)
            {
                result = null;
            }
            else
            {
                Dictionary<string, string> surface = Utils.GetSurface();
                ResultBuffer resultBuffer = poly.XData;
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
                            result = new object[]
							{
								array,
								poly
							};
                            return result;
                        }
                    }
                    result = null;
                }
            }
            return result;
        }

        public string GetSurface(Polyline poly)
        {
            Dictionary<string, string> polySurface = Utils.GetPolySurface();
            ResultBuffer xDataForApplication = poly.GetXDataForApplication("SOUTH");
            string result;
            if (xDataForApplication != null)
            {
                TypedValue[] array = xDataForApplication.AsArray();
                if (array.Length == 2)
                {
                    string text = "";
                    TypedValue[] array2 = array;
                    for (int i = 0; i < array2.Length; i++)
                    {
                        TypedValue typedValue = array2[i];
                        if (polySurface.TryGetValue(typedValue.Value.ToString(), out text))
                        {
                            result = text;
                            return result;
                        }
                    }
                }
            }
            result = "";
            return result;
        }

        public object[] GetFw_Surface(Line poly)
        {
            object[] array = new object[2];
            array[0] = "";
            Utils.polyFwDic = Utils.GetFw_PolySurface();
            string text = "";
            ResultBuffer xData = poly.XData;
            object[] result;
            if (xData != null)
            {
                TypedValue[] array2 = xData.AsArray();
                xData.Dispose();
                TypedValue[] array3 = array2;
                for (int i = 0; i < array3.Length; i++)
                {
                    TypedValue typedValue = array3[i];
                    if (Utils.polyFwDic.TryGetValue(typedValue.Value.ToString(), out text))
                    {
                        array[0] = text;
                        array[1] = poly;
                        result = array;
                        return result;
                    }
                }
            }
            result = array;
            return result;
        }

        public object[] GetFw_Surface(Polyline poly)
        {
            object[] array = new object[2];
            array[0] = "";
            Utils.polyFwDic = Utils.GetFw_PolySurface();
            string text = "";
            ResultBuffer xData = poly.XData;
            object[] result;
            if (xData != null)
            {
                TypedValue[] array2 = xData.AsArray();
                xData.Dispose();
                TypedValue[] array3 = array2;
                for (int i = 0; i < array3.Length; i++)
                {
                    TypedValue typedValue = array3[i];
                    if (Utils.polyFwDic.TryGetValue(typedValue.Value.ToString(), out text))
                    {
                        array[0] = text;
                        array[1] = poly;
                        result = array;
                        return result;
                    }
                }
            }
            result = array;
            return result;
        }

        public string GetSurface(BlockReference block)
        {
            Dictionary<string, string> surface = Utils.GetSurface();
            ResultBuffer xData = block.XData;
            string result;
            if (xData == null)
            {
                result = "";
            }
            else
            {
                TypedValue[] array = xData.AsArray();
                xData.Dispose();
                TypedValue[] array2 = array;
                for (int i = 0; i < array2.Length; i++)
                {
                    TypedValue typedValue = array2[i];
                    string text;
                    if (surface.TryGetValue(typedValue.Value.ToString(), out text))
                    {
                        result = text;
                        return result;
                    }
                }
                result = "";
            }
            return result;
        }

        public double PointAngle(Point3d startPoint, Point3d endPoint)
        {
            double x = startPoint.X;
            double y = startPoint.Y;
            double x2 = endPoint.X;
            double y2 = endPoint.Y;
            double num = x2 - x;
            double num2 = y2 - y;
            double value = num2 / num;
            double num3 = Math.Tanh(Math.Abs(value));
            return num3 * 180.0 / 3.1415926535897931;

        }

        public void Scale(Polyline3d poly, Point3d basicPoint, double scaleFactor)
        {
            Matrix3d matrix3d = Matrix3d.Scaling(scaleFactor, basicPoint);
            poly.TransformBy(matrix3d);
        }

        public Point3d GetZpoint(Point3d onePoint, Point3d twoPoint)
        {
            double num = (onePoint.X + twoPoint.X) / 2.0;
            double num2 = (onePoint.Y + twoPoint.Y) / 2.0;
            return new Point3d(num, num2, 0.0);
        }

        public Point3d GetPolyCore(Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            return this.GePolyCore(point3dCollection);
        }

        public Point3d GePolyCore(Point3dCollection northPoints)
        {
            double num = 0.0;
            double num2 = 0.0;
            foreach (Point3d point3d in northPoints)
            {
                num += point3d.X;
                num2 += point3d.Y;
            }
            return new Point3d(num / (double)northPoints.Count, num2 / (double)northPoints.Count, 0.0);
        }

        public double CalculateTwoPt(Point3d pt1, Point3d pt2)
        {
            double num = pt1.X - pt2.X;
            double num2 = pt1.Y - pt2.Y;
            return Math.Sqrt(num * num + num2 * num2);
        }

        public int PolyDiPoly(Point3dCollection pc, Point3dCollection pc2)
        {
            int num = 0;
            for (int i = 0; i < pc2.Count; i++)
            {
                double num2;
                if (i == pc2.Count - 1)
                {
                    num2 = this.PolyDistance(pc, pc2[i], pc2[0]);
                }
                else
                {
                    num2 = this.PolyDistance(pc, pc2[i], pc2[i + 1]);
                }
                if (num2 < 0.02)
                {
                    num++;
                }
            }
            return num;
        }

        public double PolyDistance(Point3dCollection pc, Point3d startPont, Point3d endPoint)
        {
            double num = 100000.0;
            foreach (Point3d point in pc)
            {
                double num2 = this.CalculateDi(point, startPont, endPoint);
                if (num > num2)
                {
                    num = num2;
                }
            }
            return num;
        }

        public double CalculateDi(Point3d point, Point3d startPoint, Point3d endPoint)
        {
            double num = this.CalculateTwoPt(startPoint, endPoint);
            double num2 = this.CalculateTwoPt(endPoint, point);
            double num3 = this.CalculateTwoPt(point, startPoint);
            double num4 = 0.5 * (num + num2 + num3);
            double num5 = Math.Sqrt(num4 * (num4 - num) * (num4 - num2) * (num4 - num3));
            return 2.0 * num5 / num3;
        }

        public double GetDisantaceMax(int[] pointIndex, Point3dCollection points, Point3dCollection fourPoints)
        {
            Point3d pt = points[pointIndex[0]];
            Point3d pt2 = points[pointIndex[1]];
            Point3d pt3 = points[pointIndex[2]];
            Point3d pt4 = points[pointIndex[3]];
            Point3d pt5 = fourPoints[0];
            Point3d pt6 = fourPoints[1];
            Point3d pt7 = fourPoints[2];
            Point3d pt8 = fourPoints[3];
            if (this.CalculateTwoPt(pt, pt5) > this.CalculateTwoPt(pt, pt6))
            {
                pt = points[pointIndex[1]];
                pt2 = points[pointIndex[2]];
                pt3 = points[pointIndex[3]];
                pt4 = points[pointIndex[0]];
            }
            double num = this.CalculateTwoPt(pt, pt5);
            double num2 = this.CalculateTwoPt(pt2, pt6);
            double num3 = this.CalculateTwoPt(pt3, pt7);
            double num4 = this.CalculateTwoPt(pt4, pt8);
            double[] array = new double[]
			{
				num,
				num2,
				num3,
				num4
			};
            int doubleMaxIndex = this.GetDoubleMaxIndex(array);
            return array[doubleMaxIndex];
        }

        public int GetDoubleMaxIndex(double[] d)
        {
            double num = d[0];
            int result = 0;
            for (int i = 1; i < d.Length; i++)
            {
                if (num < d[i])
                {
                    num = d[i];
                    result = i;
                }
            }
            return result;
        }

        public double PolyDistanceMax(Point3dCollection pc, Point3d startPont, Point3d endPoint)
        {
            double num = 0.0;
            foreach (Point3d point in pc)
            {
                double num2 = this.CalculateDi(point, startPont, endPoint);
                if (num < num2)
                {
                    num = num2;
                }
            }
            return num;
        }

        public Point3dCollection CheckRePoint(Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            return this.CheckRePoint(point3dCollection);
        }

        public Point3dCollection CheckRePoint(Point3dCollection points)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            Point3dCollection point3dCollection2 = new Point3dCollection();
            foreach (Point3d point3d in points)
            {
                if (!this.RePoint(point3d, point3dCollection))
                {
                    point3dCollection.Add(point3d);
                }
                else
                {
                    point3dCollection2.Add(point3d);
                }
            }
            return point3dCollection2;
        }

        public Point3dCollection RePoint(Point3dCollection points)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            foreach (Point3d point3d in points)
            {
                if (!this.RePoint(point3d, point3dCollection))
                {
                    point3dCollection.Add(point3d);
                }
            }
            return point3dCollection;
        }

        public bool RePoint(Point3d point, Point3dCollection newNorthPoints)
        {
            bool result;
            for (int i = 0; i < newNorthPoints.Count; i++)
            {
                if (Math.Abs(point.X - newNorthPoints[i].X) < 0.001 && Math.Abs(point.Y - newNorthPoints[i].Y) < 0.001)
                {
                    result = true;
                    return result;
                }
            }
            result = false;
            return result;
        }

        public string SelectSur(Polyline3d poly)
        {
            DBObjectCollection corssEntityCollection = this.GetCorssEntityCollection(poly);
            string result;
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline poly2 = (Polyline)dBObject;
                    string surface = this.GetSurface(poly2);
                    if (surface != "")
                    {
                        result = surface;
                        return result;
                    }
                }
                if (dBObject is BlockReference)
                {
                    BlockReference block = (BlockReference)dBObject;
                    string surface = this.GetSurface(block);
                    if (surface != "")
                    {
                        result = surface;
                        return result;
                    }
                }
                if (dBObject is DBText)
                {
                    DBText dBText = (DBText)dBObject;
                    string text = "";
                    if (Utils.txDic.TryGetValue(dBText.TextString, out text))
                    {
                        result = text;
                        return result;
                    }
                }
            }
            result = "";
            return result;
        }

        public object[] selectJzxSur(Point3d pt1, Point3d pt2, double di, Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            Polyline polyline = new Polyline();
            polyline.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0.0, 0.0, 0.0);
            polyline.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0.0, 0.0, 0.0);
            DBObjectCollection offsetCurves = polyline.GetOffsetCurves(0.05);
            foreach (DBObject dBObject in offsetCurves)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int i = 0; i < point3dCollection2.Count; i++)
                    {
                        point3dCollection.Add(point3dCollection2[i]);
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
                    for (int i = point3dCollection2.Count - 1; i >= 0; i--)
                    {
                        point3dCollection.Add(point3dCollection2[i]);
                    }
                }
            }
            Polyline3d poly2 = this.CreatePolye(point3dCollection, true);
            this.Scale(poly2, this.GetZpoint(point3dCollection[0], point3dCollection[1]), 0.99);
            object[] surfacePoly3d = this.GetSurfacePoly3d(poly2, poly);
            if (surfacePoly3d == null && FormService.scheme != 0)
            {
                string text = this.SelectSur(poly2);
            }
            return surfacePoly3d;
        }

        public int SelectAreaMax(int[][] max, Point3dCollection pc)
        {
            double num = 0.0;
            int result = 0;
            for (int i = 0; i < max.Length; i++)
            {
                Point3dCollection pc2 = this.SelectPoints(max[i], pc);
                double num2 = this.PointsArea(pc2);
                if (num2 > num)
                {
                    num = num2;
                    result = i;
                }
            }
            return result;
        }

        public int SelectPc1PC2(int[][] max, Point3dCollection pc)
        {
            Point3dCollection pc2 = this.SelectPoints(max[0], pc);
            Point3dCollection pc3 = this.SelectPoints(max[1], pc);
            double num = this.PointsArea(pc2);
            double num2 = this.PointsArea(pc3);
            double num3 = num - num2;
            int result;
            if (Math.Abs(num3) / num < 0.01 || num3 > 0.0)
            {
                result = 0;
            }
            else
            {
                result = 1;
            }
            return result;
        }

        public double PointsArea(Point3dCollection pc)
        {
            Polyline3d polyline3d = new Polyline3d(0, pc, true);
            return polyline3d.Area;
        }

        public Point3dCollection SelectPoints(int[] max, Point3dCollection pc)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            for (int i = 0; i < max.Length; i++)
            {
                point3dCollection.Add(pc[max[i]]);
            }
            return point3dCollection;
        }

        //日期转换
        public static DateTime DateFormat(String dateTime)
        {
            DateTime dt;

            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();

            dtFormat.ShortDatePattern = "yyyy年MM月dd日  HH:mm:ss";

            dt = Convert.ToDateTime(dateTime, dtFormat);
            return dt;
        }

        public Polyline3d PolyOffset(object polyO, double inner, double outer)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            Polyline polyline = new Polyline();
            Point3dCollection point3dCollection2 = new Point3dCollection();
            if (polyO is Polyline)
            {
                Polyline polyline2 = (Polyline)polyO;
                polyline2.GetStretchPoints(point3dCollection2);
            }
            else if (polyO is Polyline3d)
            {
                Polyline3d polyline3d = (Polyline3d)polyO;
                polyline3d.GetStretchPoints(point3dCollection2);
            }
            for (int i = 0; i < point3dCollection2.Count; i++)
            {
                Point3d point3d = point3dCollection2[i];
                polyline.AddVertexAt(i, new Point2d(point3d.X, point3d.Y), 0.0, 0.0, 0.0);
            }
            if (outer == 0.0)
            {
                for (int i = point3dCollection2.Count - 1; i >= 0; i--)
                {
                    Point3d point3d = point3dCollection2[i];
                    point3dCollection.Add(point3d);
                }
            }
            if (inner == 0.0)
            {
                for (int i = 0; i < point3dCollection2.Count; i++)
                {
                    Point3d point3d = point3dCollection2[i];
                    point3dCollection.Add(point3d);
                }
            }
            if (inner > 0.0)
            {
                DBObjectCollection offsetCurves = polyline.GetOffsetCurves(inner);
                foreach (DBObject dBObject in offsetCurves)
                {
                    if (dBObject is Polyline)
                    {
                        Polyline polyline3 = (Polyline)dBObject;
                        Point3dCollection point3dCollection3 = new Point3dCollection();
                        polyline3.GetStretchPoints(point3dCollection3);
                        for (int j = 0; j < point3dCollection3.Count; j++)
                        {
                            point3dCollection.Add(point3dCollection3[j]);
                        }
                    }
                }
            }
            if (outer > 0.0)
            {
                DBObjectCollection offsetCurves2 = polyline.GetOffsetCurves(-outer);
                foreach (DBObject dBObject in offsetCurves2)
                {
                    if (dBObject is Polyline)
                    {
                        Polyline polyline3 = (Polyline)dBObject;
                        Point3dCollection point3dCollection3 = new Point3dCollection();
                        polyline3.GetStretchPoints(point3dCollection3);
                        for (int j = point3dCollection3.Count - 1; j >= 0; j--)
                        {
                            point3dCollection.Add(point3dCollection3[j]);
                        }
                    }
                }
            }
            Polyline3d polyline3d2 = this.CreatePolye(point3dCollection, true);
            this.Scale(polyline3d2, this.GetZpoint(point3dCollection[0], point3dCollection[1]), 0.99);
            this.CreatePolye(point3dCollection, true);
            return polyline3d2;
        }

        public string selectFw_JzxSur(Point3d pt1, Point3d pt2, Polyline poly)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            Polyline polyline = new Polyline();
            polyline.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0.0, 0.0, 0.0);
            polyline.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0.0, 0.0, 0.0);
            DBObjectCollection offsetCurves = polyline.GetOffsetCurves(0.02);
            foreach (DBObject dBObject in offsetCurves)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int i = 0; i < point3dCollection2.Count; i++)
                    {
                        point3dCollection.Add(point3dCollection2[i]);
                    }
                }
            }
            DBObjectCollection offsetCurves2 = polyline.GetOffsetCurves(-0.02);
            foreach (DBObject dBObject in offsetCurves2)
            {
                if (dBObject is Polyline)
                {
                    Polyline polyline2 = (Polyline)dBObject;
                    Point3dCollection point3dCollection2 = new Point3dCollection();
                    polyline2.GetStretchPoints(point3dCollection2);
                    for (int i = point3dCollection2.Count - 1; i >= 0; i--)
                    {
                        point3dCollection.Add(point3dCollection2[i]);
                    }
                }
            }
            Polyline3d poly2 = this.CreatePolye(point3dCollection, true);
            this.Scale(poly2, this.GetZpoint(point3dCollection[0], point3dCollection[1]), 0.9);
            object[] array = this.SelectBS_Sur(poly2);
            string text = (string)array[0];
            if (text.Equals(""))
            {
                text = "界址线";
            }
            string str = "外";
            TypedValue[] array2 = null;
            object[] surfacePoly3d = this.GetSurfacePoly3d(poly2, poly);
            if (surfacePoly3d != null)
            {
                array2 = (TypedValue[])surfacePoly3d[0];
            }
            if (array2 != null && array2.Length > 2)
            {
                if (array2[2].Value.ToString().Contains("J"))
                {
                    if (array[1] == null)
                    {
                        str = "中";
                    }
                    else if (array[2] != null && array[0].ToString().Equals(array[2].ToString()))
                    {
                        str = "中";
                    }
                    else
                    {
                        Autodesk.AutoCAD.DatabaseServices.Region region = this.RegionOpreation(poly, (Polyline)array[1]);
                        if (region == null || region.Area> 0.5)
                        {
                            str = "外";
                        }
                        else
                        {
                            str = "内";
                        }
                    }
                }
            }
            return text + "_" + str;
        }

        public object[] SelectBS_Sur(Polyline3d poly)
        {
            DBObjectCollection corssEntityCollection = this.GetCorssEntityCollection(poly);
            object[] array = new object[3];
            array[0] = "";
            object[] result;
            foreach (DBObject dBObject in corssEntityCollection)
            {
                if (dBObject is Polyline)
                {
                    Polyline poly2 = (Polyline)dBObject;
                    object[] fw_Surface = this.GetFw_Surface(poly2);
                    string text = (string)fw_Surface[0];
                    if (!text.Equals(""))
                    {
                        if (!array[0].Equals(""))
                        {
                            array[2] = text;
                            result = array;
                            return result;
                        }
                        array[0] = fw_Surface[0];
                        array[1] = fw_Surface[1];
                    }
                }
                else if (dBObject is Line)
                {
                    Line poly3 = (Line)dBObject;
                    object[] fw_Surface = this.GetFw_Surface(poly3);
                    string text = (string)fw_Surface[0];
                    if (!text.Equals(""))
                    {
                        if (!array[0].Equals(""))
                        {
                            array[2] = text;
                            result = array;
                            return result;
                        }
                        array[0] = fw_Surface[0];
                        array[1] = fw_Surface[1];
                    }
                }
            }
            result = array;
            return result;
        }

        public string getWindowText(Point3d pt1, Point3d pt2)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "TEXT")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            Point3dCollection point3dCollection = new Point3dCollection();
            Point3d point3d = new Point3d(pt2.X, pt1.Y, 0.0);
            Point3d point3d2 = new Point3d(pt1.X, pt2.Y, 0.0);
            point3dCollection.Add(pt1);
            point3dCollection.Add(point3d2);
            point3dCollection.Add(pt2);
            point3dCollection.Add(point3d);
            Polyline3d polyline3d = this.CreatePolye(point3dCollection, true);
            double area = polyline3d.Area;
            Point3dCollection point3dCollection2 = new Point3dCollection();
            polyline3d.GetStretchPoints(point3dCollection2);
            PromptSelectionResult ents = editor.SelectCrossingPolygon(point3dCollection2);
            DBObjectCollection entityCollection = this.GetEntityCollection(ents);
            string text = "";
            foreach (Entity entity in entityCollection)
            {
                if (entity is DBText)
                {
                    DBText dBText = entity as DBText;
                    if (dBText != null)
                    {
                        if (!dBText.TextString.Equals(""))
                        {
                        }
                        text += dBText.TextString;
                    }
                }
            }
            return text;
        }

        public List<string> getWindowText(Polyline poly)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "TEXT")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            point3dCollection = this.RePoint(point3dCollection);
            PromptSelectionResult ents = editor.SelectCrossingPolygon(point3dCollection);
            DBObjectCollection entityCollection = this.GetEntityCollection(ents);
            List<string> list = new List<string>();
            foreach (Entity entity in entityCollection)
            {
                if (entity is DBText)
                {
                    DBText dBText = entity as DBText;
                    if (dBText != null)
                    {
                        if (!dBText.TextString.Equals(""))
                        {
                            list.Add(dBText.TextString);
                        }
                    }
                }
            }
            return list;
        }

        public void CreateCircle(Point3dCollection pc, int colorIndex, string layerName)
        {
            foreach (Point3d center in pc)
            {
                this.CreateCircle(center, colorIndex, layerName);
            }
        }

        public void CreateCircle(Point3d center, int colorIndex, string layerName)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, 0);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                Circle circle = new Circle(center, Vector3d.ZAxis, 2.0);
                circle.ColorIndex=(colorIndex);
                Utils.CreateLayer(layerName);
                circle.Layer=(layerName);
                blockTableRecord.AppendEntity(circle);
                transaction.AddNewlyCreatedDBObject(circle, true);
                transaction.Commit();
            }
        }

             //创建文字样式
            public ObjectId CreateSytle1(string stylename, double xScale)
        {
            ObjectId textstyleid = new ObjectId();
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                TextStyleTable st = trans.GetObject(db.TextStyleTableId, OpenMode.ForWrite) as TextStyleTable;
                //string stylename = "gbc";
                TextStyleTableRecord str = new TextStyleTableRecord();

                        //查询是否包含gbc样式，若果不包含就创建该文字样式，包含就取出。
                if (!st.Has(stylename))
                {
                            str.Name = stylename;
                            str.XScale = (xScale);
                    str.FileName = "txt";
                    str.BigFontFileName = "gbcbig";
                    
                  
                }
                else
                {
                    foreach (ObjectId id in st)
                    {
                                TextStyleTableRecord stlye = trans.GetObject(id, OpenMode.ForWrite) as TextStyleTableRecord;
                                if (stlye.Name == stylename)
                        {
                            return id;
                        }
                    }
                }
                textstyleid = st.Add(str);
                trans.AddNewlyCreatedDBObject(str, true);
                trans.Commit();
            }
            return textstyleid;
        }



        public ObjectId CreateSytle(string fileName, double xScale)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            ObjectId result;
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                TextStyleTable textStyleTable = (TextStyleTable)transaction.GetObject(workingDatabase.TextStyleTableId, OpenMode.ForWrite);
                TextStyleTableRecord textStyleTableRecord = new TextStyleTableRecord();
                textStyleTableRecord.XScale = (xScale);

                ObjectId objectId = textStyleTable.Add(textStyleTableRecord);
                transaction.AddNewlyCreatedDBObject(textStyleTableRecord, true);
                workingDatabase.Textstyle = (objectId);
                transaction.Commit();
                result = objectId;
            }
            return result;
        }

        public void AddDBText(DBText dbtext)
        {
            if (dbtext.TextString != null && !dbtext.TextString.Equals(""))
            {
                Database workingDatabase = HostApplicationServices.WorkingDatabase;
                using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId,OpenMode.ForWrite);
                    BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    blockTableRecord.AppendEntity(dbtext);
                    transaction.AddNewlyCreatedDBObject(dbtext, true);
                    transaction.Commit();
                }
            }
        }

        public DBText CreateDBText(string textString, Point3d position, double height, ObjectId textStyle)
        {
            DBText dBText = new DBText();
          
            dBText.Position=(position);
           
            dBText.Height=(height);
            if (textString == null)
            {
                textString = "";
            }
            dBText.TextString=(textString);
            dBText.TextStyle=(textStyle);
            dBText.WidthFactor = 0.8;
            return dBText;
        }

        public DBObjectCollection SelectDBObjectCollection(string message, TypedValue[] types)
        {
            DBObjectCollection dBObjectCollection = new DBObjectCollection();
            Editor editor = this.GetEditor();
            PromptSelectionOptions promptSelectionOptions = new PromptSelectionOptions();
            promptSelectionOptions.MessageForAdding=(message);
            PromptSelectionResult selection;
            if (types != null)
            {
                SelectionFilter selectionFilter = new SelectionFilter(types);
                selection = editor.GetSelection(promptSelectionOptions, selectionFilter);
            }
            else
            {
                selection = editor.GetSelection(promptSelectionOptions);
            }
            DBObjectCollection result;
            if (selection.Status == PromptStatus.Cancel)
            {
                result = dBObjectCollection;
            }
            else
            {
                dBObjectCollection = this.GetEntityCollection(selection);
                if (dBObjectCollection.Count == 0)
                {
                    editor.WriteMessage("\n你没有选择对象");
                }
                result = dBObjectCollection;
            }
            return result;
        }

        public void deleteLayerSur(string layerName)
        {
            Utils utils = new Utils();
            Editor editor = utils.GetEditor();
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                TypedValue[] array = new TypedValue[1];
                array.SetValue(new TypedValue(8, layerName), 0);
                SelectionFilter selectionFilter = new SelectionFilter(array);
                PromptSelectionResult promptSelectionResult = editor.SelectAll(selectionFilter);
                if (promptSelectionResult.Status != PromptStatus.OK) { return; }
                {
                    SelectionSet value = promptSelectionResult.Value;
                    ObjectId[] objectIds = value.GetObjectIds();
                    for (int i = 0; i < objectIds.Length; i++)
                    {
                        ObjectId objectId = objectIds[i];
                        Entity entity = transaction.GetObject(objectId, OpenMode.ForWrite) as Entity;
                        entity.Erase();
                    }
                    transaction.Commit();
                }
            }
        }

        public Autodesk.AutoCAD.DatabaseServices.Region RegionOpreation(Autodesk.AutoCAD.DatabaseServices.Region r1, Autodesk.AutoCAD.DatabaseServices.Region r2)
        {
            r1.BooleanOperation(BooleanOperationType.BoolIntersect, r2);

            return r1;
        }

        public Autodesk.AutoCAD.DatabaseServices.Region RegionOpreation(Polyline pl1, Polyline pl2)
        {
            PolyService polyService = new PolyService();
            if (pl1.Area== 0.0)
            {
                Polyline3d poly3d = this.PolyOffset(pl1, 0.0, 5.0);
                pl1 = polyService.Poly3dToPolyLine(poly3d);
            }
            else
            {
                Point3dCollection point3dCollection = new Point3dCollection();
                pl1.GetStretchPoints(point3dCollection);
                pl1 = polyService.Poly3dToPolyLine(this.CreatePolye(point3dCollection, true));
            }
            if (pl2.Area== 0.0)
            {
                Polyline3d poly3d = this.PolyOffset(pl2, 0.0, 5.0);
                pl2 = polyService.Poly3dToPolyLine(poly3d);
            }
            else
            {
                Point3dCollection point3dCollection2 = new Point3dCollection();
                pl2.GetStretchPoints(point3dCollection2);
                pl2 = polyService.Poly3dToPolyLine(this.CreatePolye(point3dCollection2, true));
            }
            pl1.Closed=(true);
            pl2.Closed=(true);
            DBObjectCollection dBObjectCollection = new DBObjectCollection();
            dBObjectCollection.Add(pl1);
            dBObjectCollection.Add(pl2);
            Region region = null;
            Region region2 = null;
            Region result;
            try
            {
                DBObjectCollection dBObjectCollection2 = Region.CreateFromCurves(dBObjectCollection);
                int num = 0;
                foreach (DBObject dBObject in dBObjectCollection2)
                {
                    if (num == 1)
                    {
                        region2 = (dBObject as Region);
                    }
                    else if (num == 0)
                    {
                        region = (dBObject as Region);
                    }
                    num++;
                }
               
                region.BooleanOperation(BooleanOperationType.BoolIntersect, region2);
            }
            catch
            {
                result = null;
                return result;
            }
            result = region;
            return result;
        }

        public static string ListToStr(string[] array, string connect)
        {
            string text = "";
            for (int i = 0; i < array.Length; i++)
            {
                if (i != 0 && i != array.Length - 1)
                {
                    text += connect;
                }
                text += array[i];
            }
            return text;
        }

        public void ListToText(List<string> list, string fileName)
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

        public bool WClone(DBObjectCollection dbo, string FileName)
        {
            ObjectIdCollection dBObjectCollectionIds = this.GetDBObjectCollectionIds(dbo);
            return this.WClone(dBObjectCollectionIds, FileName);
        }

        public ObjectIdCollection GetDBObjectCollectionIds(DBObjectCollection dbo)
        {
            ObjectIdCollection objectIdCollection = new ObjectIdCollection();
            foreach (DBObject dBObject in dbo)
            {
                objectIdCollection.Add(dBObject.ObjectId);
            }
            return objectIdCollection;
        }

        public Boolean WClone(ObjectIdCollection idc, String savePath)
        {
            if (idc == null || idc.Count == 0)
            {
                return false;
            }
            Database ndb = new Database(true, true);
            ObjectId idbtr = new ObjectId();
            Database db = idc[0].Database;
            IdMapping map = new IdMapping();
            using (Transaction trans = ndb.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(ndb.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                idbtr = modelSpace.ObjectId;
                trans.Commit();

            }

           db.WblockCloneObjects(idc, idbtr, map, DuplicateRecordCloning.Replace, false);
            //db.WblockCloneObjects(idc, idbtr, DuplicateRecordCloning.Replace, false);
            ndb.SaveAs(savePath, DwgVersion.Current);

            ZoomExtens(savePath);
            return true;


        }

        public static void ZoomExtens(string fileName)
        {
            DocumentCollection documentManager = Application.DocumentManager;
            Document document = documentManager.Open(fileName, false);
            DocumentLock documentLock = document.LockDocument();
            System.Drawing.Size size = new System.Drawing.Size(150, 150);
            Editor editor = document.Editor;
            document.Window.Size=size;
            Utils.ZoomExtens(editor);
            documentLock.Dispose();
            document.CloseAndSave(fileName);
        }

        public static void ZoomExtens(Editor ed)
        {
            Database database = ed.Document.Database;
            database.UpdateExt(true);
            if (database.Extmax.X < database.Extmin.X)
            {
                Plane plane = new Plane();
                Point3d pt = new Point3d(plane, database.Limmin);
                Point3d pt2 = new Point3d(plane, database.Limmax);
                Utils.ZoomWindow(ed, pt, pt2);
            }
            else
            {
                Utils.ZoomWindow(ed, database.Extmin, database.Extmax);
            }
        }
        /* */
        public static void ZoomWindow(Editor ed, Point3d pt1, Point3d pt2)
        {
            using (Line line = new Line(pt1, pt2))
            {
                Extents3d extents3d = new Extents3d(line.GeometricExtents.MinPoint, line.GeometricExtents.MaxPoint);
                Point2d point2d = new Point2d(extents3d.MinPoint.X, extents3d.MinPoint.Y);
                Point2d point2d2 = new Point2d(extents3d.MaxPoint.X, extents3d.MaxPoint.Y);
                ViewTableRecord currentView = ed.GetCurrentView();
                currentView.CenterPoint=(point2d + (point2d2 - point2d) / 2.0);
                currentView.Height=(point2d2.Y - point2d.Y);
                currentView.Width=(point2d2.X - point2d.X);
                ed.SetCurrentView(currentView);
            }
        }
        /*   
        /// <summary>
        /// 06cad采用
        /// </summary>
        /// <param name="ed"></param>
        /// <param name="pt1"></param>
        /// <param name="pt2"></param>
        public static void ZoomWindow(Editor ed, Point3d pt1, Point3d pt2)
        {
            //创建一临时的直线用于获取两点表示的范围
            using (Line line = new Line(pt1, pt2))
            {
                //获取两点表示的范围

                Extents3d extents = new Extents3d(line.GeomExtents.MinPoint, line.GeomExtents.MaxPoint);
                //获取范围内的最小值点及最大值点
                Point2d minPt = new Point2d(extents.MinPoint.X, extents.MinPoint.Y);
                Point2d maxPt = new Point2d(extents.MaxPoint.X, extents.MaxPoint.Y);

                //得到当前视图
                // ViewTableRecord view = ed.GetCurrentView();
                ed.UpdateScreen();
            }
        }
* */
        public Document GetNewDocument()
        {
            DocumentCollection documentManager = Application.DocumentManager;
            return documentManager.Add("ACAD.DWT");
        }

        internal int[] FindAreaMax(Point3dCollection points, int[] pointIndex1)
        {
            int[] array = this.CopyArray<int>(pointIndex1, 4);
            int[] array2 = this.CopyArray<int>(array, 4);
            double num = this.SelectPointsArea(points, array);
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < points.Count; j++)
                {
                    array2[i] = j;
                    double num2 = this.SelectPointsArea(points, array2);
                    if (num < num2)
                    {
                        array[i] = j;
                        num = num2;
                    }
                }
                array2 = this.CopyArray<int>(array, 4);
            }
            array = this.IfXmax(points, array);
            num = this.SelectPointsArea(points, array);
            return array;
        }

        public int[] IfXmax(Point3dCollection points, int[] pointIndex)
        {
            int[] result;
            if (points[pointIndex[0]].X < points[pointIndex[1]].X)
            {
                result = pointIndex;
            }
            else
            {
                int[] array = new int[]
				{
					pointIndex[0],
					0,
					0,
					0,
					pointIndex[0]
				};
                array[1] = pointIndex[3];
                array[2] = pointIndex[2];
                array[3] = pointIndex[1];
                result = array;
            }
            return result;
        }

        private T[] CopyArray<T>(T[] array, int length)
        {
            T[] array2 = new T[length];
            for (int i = 0; i < length; i++)
            {
                array2[i] = array[i];
            }
            return array2;
        }

        private T[] CopyArray<T>(T[] array)
        {
            int num = array.Length;
            T[] array2 = new T[num];
            for (int i = 0; i < num; i++)
            {
                array2[i] = array[i];
            }
            return array2;
        }

        public double SelectPointsArea(Point3dCollection points, int[] pointIndex)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            for (int i = 0; i < pointIndex.Length; i++)
            {
                int num = pointIndex[i];
                point3dCollection.Add(points[num]);
            }
            return this.PointsArea(point3dCollection);
        }

        internal int[] GetMaxArea(Point3dCollection points, int[] pointIndex, int[] pointIndex2)
        {
            double num = this.SelectPointsArea(points, pointIndex);
            double num2 = this.SelectPointsArea(points, pointIndex2);
            double num3 = num - num2;
            int[] result;
            if (Math.Abs(num3) / num < 0.03 || num3 > 0.0)
            {
                result = pointIndex;
            }
            else
            {
                pointIndex = new int[]
				{
					pointIndex2[0],
					0,
					0,
					0,
					pointIndex2[0]
				};
                pointIndex[1] = pointIndex2[1];
                pointIndex[2] = pointIndex2[2];
                pointIndex[3] = pointIndex2[3];
                result = pointIndex;
            }
            return result;
        }

        internal static bool IsStrNull(string fileName)
        {
            return fileName == null || fileName.Trim().Equals("");
        }

        internal Transaction GetTransaction()
        {
            return HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction();
        }

        internal Polyline3d Scale(Polyline pl,double big)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            pl.GetStretchPoints(point3dCollection);
            Polyline3d polyline3d = CreatePolye(point3dCollection, true);
            Point3d core = GetPolyCore(pl);
            Scale(polyline3d, core, big);
             return polyline3d;
        }
    }
}
