using System;
using System.Collections.Generic;
using System.Text;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using YanBo_CG;
namespace YanBo
{
    public class CADUtils
    {
        private static Utils utils = new Utils();

        internal static void MergeGeometrys(DBObjectCollection dbc)
        {
            if (dbc.Count > 1 && dbc[0] is Polyline)
            {
                Region r = null;
                TypedValue[] tv = null;
                using (Transaction trans = utils.GetDabase().TransactionManager.StartTransaction())
                {

                    for (int a = 0; a < dbc.Count; a++)
                    {
                        if (r == null)
                        {
                            r = CreateRegion(dbc[a]);

                        }
                        else
                        {
                            r.BooleanOperation(BooleanOperationType.BoolUnite, CreateRegion(dbc[a]));
                        }
                        ResultBuffer rb = dbc[a].GetXDataForApplication("SOUTH");

                        if (rb != null)
                        {
                            if (tv == null)
                            {
                                tv = rb.AsArray();
                            }
                            else if (rb.AsArray().Length > 4 && rb.AsArray()[3] != null && !rb.AsArray()[3].Value.ToString().Trim().Equals(""))
                            {
                                string s = tv[3].Value.ToString() + "、" + rb.AsArray()[3].Value.ToString();
                                tv[3] = new TypedValue((int)DxfCode.ExtendedDataAsciiString, s.Replace("（户）、", "、"));
                            }
                        }


                        DeleteEntity(trans, dbc[a].Id);

                    }
                    Point3dCollection pc = new Point3dCollection();
                    r.GetStretchPoints(pc);
                    pc = Adjustment(pc);
                    Polyline plSrc = (Polyline)dbc[0];

                    Polyline pl = PCToPolyline(pc, plSrc.GetStartWidthAt(0), true);
                    CopyAttribute(plSrc, pl);
                    if (tv != null)
                    {
                        pl.XData = new ResultBuffer(tv);
                    }
                    SaveEntity(trans, pl);
                    trans.Commit();
                }

            }

        }
        public static void Adjustment(DBObjectCollection dbc)
        {
            using (Transaction trans = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject db in dbc)
                {
                    if (db is Polyline)
                    {

                        Polyline pl = (Polyline)db;
                        if (pl.Closed)//封闭的才转
                        {
                            Polyline p = Adjustment(pl);
                            SaveEntity(trans, p);
                            DeleteEntity(trans, db.ObjectId);
                        }

                    }
                }
                trans.Commit();
            }
        }
        private static Polyline Adjustment(Polyline pl)
        {

            Point3dCollection pc = new Point3dCollection();
            pl.GetStretchPoints(pc);
            pc = Adjustment(pc);
            Polyline p = PCToPolyline(pc, pl.GetStartWidthAt(0), pl.Closed);
            CopyAttribute(pl, p);
            return p;

        }
        private static Point3dCollection Adjustment(Point3dCollection pc)
        {
            double minX = FinMinX(pc);
            double maxY = FinMaxY(pc);
            Point3d point = new Point3d(minX, maxY, 0);
            double di = utils.CalculateTwoPt(point, pc[0]);
            int count = pc.Count;
            int index = 0;
            double compare;
            for (int a = 1; a < count; a++)
            {
                compare = utils.CalculateTwoPt(point, pc[a]);
                if (di > compare)
                {
                    di = compare;
                    index = a;
                }
            }
            pc = RecombinationPoints(pc, index);
            if (pc[0].X > pc[pc.Count - 1].X)
            {
                return pc;
            }
            else
            {
                Point3dCollection result = new Point3dCollection();
                result.Add(pc[0]);
                for (int a = pc.Count - 1; a > 0; a--)
                {
                    result.Add(pc[a]);
                }
                return result;
            }

        }

        private static Point3dCollection RecombinationPoints(Point3dCollection pc, int index)
        {
            Point3dCollection reuslt = new Point3dCollection();
            for (int a = index; a < pc.Count; a++)
            {
                reuslt.Add(pc[a]);
            }
            for (int a = 0; a < index; a++)
            {
                reuslt.Add(pc[a]);
            }
            return reuslt;
        }

        private static double FinMaxY(Point3dCollection pc)
        {
            double y = pc[0].Y;
            for (int a = 1; a < pc.Count; a++)
            {
                if (y < pc[a].Y)
                {
                    y = pc[a].Y;
                }
            }
            return y;
        }

        private static double FinMinX(Point3dCollection pc)
        {
            double x = pc[0].X;
            for (int a = 1; a < pc.Count; a++)
            {
                if (x > pc[a].X)
                {
                    x = pc[a].X;
                }
            }
            return x;
        }
        public static void SaveEntity(Transaction trans, Entity entity)
        {

            //获取块表和模型空间
            BlockTable bt = (BlockTable)(trans.GetObject(utils.GetDabase().BlockTableId, OpenMode.ForWrite));
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            btr.AppendEntity(entity);
            trans.AddNewlyCreatedDBObject(entity, true);

        }
        public static void DeleteEntity(Transaction trans, ObjectId id)
        {
            Entity hatchobj = trans.GetObject(id, OpenMode.ForWrite) as Entity;
            hatchobj.Erase();//删除  
        }
        public static void CopyAttribute(Entity src, Entity desc)
        {
            desc.XData = src.XData;
            desc.Color = src.Color;
            desc.Linetype = src.Linetype;
            desc.LineWeight = src.LineWeight;
            desc.Layer = src.Layer;
        }
        /// <summary>
        /// 三维点集合转二维点集合
        /// </summary>
        /// <param name="pc"></param>
        /// <returns></returns>
        public static Point2dCollection Point3dCollectionTo2d(Point3dCollection pc)
        {
            Point2dCollection pc2 = new Point2dCollection();
            foreach (Point3d pt in pc)
            {
                pc2.Add(new Point2d(pt.X, pt.Y));
            }
            return pc2;
        }

        public static Polyline PCToPolyline(Point2dCollection pc, double width, bool isClosed)
        {
            Polyline poly = new Polyline();
            for (int a = 0; a < pc.Count; a++)
            {
                poly.AddVertexAt(a, pc[a], 0, width, width);

            }
            poly.Closed = isClosed;
            return poly;
        }
        public static Polyline PCToPolyline(Point3dCollection pc, double width, bool isClosed)
        {
            return PCToPolyline(Point3dCollectionTo2d(pc), width, isClosed);
        }




        private static bool IsOnePoint(Point3d pt1, Point3d pt2)
        {

            if (Math.Abs(pt1.X - pt2.X) < 0.0001 && Math.Abs(pt1.Y - pt2.Y) < 0.0001)
            {
                Console.WriteLine(pt1.X + "::" + pt1.Y);
                return true;
            }
            return false;
        }
        public static Region CreateRegion(DBObject db)
        {
            Polyline pl = null;
            if (db is Polyline)
            {
                Polyline p = (Polyline)db;
                Point3dCollection pc = new Point3dCollection();

                p.GetStretchPoints(pc);
                Point3d s = pc[0];
                Point3d e = pc[pc.Count - 1];
                if (s.X != e.X && s.Y != e.Y)
                {
                    pc.Add(s);
                }
                pl = PCToPolyline(pc, p.GetStartWidthAt(0), p.Closed);
            }
            else
            {
                return null;
            }

            DBObjectCollection dbc = new DBObjectCollection();
            if (pl != null)
            {
                dbc.Add(pl);
            }
            else
            {
                dbc.Add(db);
            }
            try
            {
                DBObjectCollection res = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(dbc);
                return (Region)res[0];
            }
            catch
            {
                return null;
            }

        }

        public static Region CreateRegion1(Polyline pl)
        {
            DBObjectCollection dbc = new DBObjectCollection();
            if (pl != null)
            {
                dbc.Add(pl);
            }
            else
            {
                return null;
            }
            try
            {
                DBObjectCollection res = Autodesk.AutoCAD.DatabaseServices.Region.CreateFromCurves(dbc);
                return (Region)res[0];
            }
            catch
            {
                return null;
            }
        }
        public static void SetText(DBText dbText, string text, string fontName, double height, double width)
        {
            dbText.TextString = text;
            dbText.TextStyle = CADUtils.GetTextStyle(fontName);
            dbText.Height = height;
            dbText.WidthFactor = width;
        }
        public static ObjectId GetTextStyle(string name)
        {
            Database dbH = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = dbH.TransactionManager.StartTransaction())
            {
                TextStyleTable TST = (TextStyleTable)trans.GetObject(dbH.TextStyleTableId, OpenMode.ForWrite);
                ObjectId id = GetIdFromSymbolTable(TST, name);
                return id;
            }
        }
        public static ObjectId AddTextStyle(string name, string smallfont, string bigfont, double height, double xscale)
        {
            Database dbH = HostApplicationServices.WorkingDatabase;

            using (Transaction trans = dbH.TransactionManager.StartTransaction())
            {
                TextStyleTable TST = (TextStyleTable)trans.GetObject(dbH.TextStyleTableId, OpenMode.ForWrite);
                ObjectId id = GetIdFromSymbolTable(TST, name);
                if (id == ObjectId.Null)
                {
                    TextStyleTableRecord TSTR = new TextStyleTableRecord();
                    TSTR.Name = name;
                    TSTR.FileName = smallfont;
                    TSTR.BigFontFileName = bigfont;
                    TSTR.TextSize = height;
                    TSTR.XScale = xscale;
                    TST.UpgradeOpen();
                    id = TST.Add(TSTR);
                    trans.AddNewlyCreatedDBObject(TSTR, true);
                }
                return id;
            }
        }
        //取得符号表的Id
        public static ObjectId GetIdFromSymbolTable(SymbolTable st, string key)
        {
            Database dbH = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = dbH.TransactionManager.StartTransaction())
            {
                if (st.Has(key))
                {
                    ObjectId idres = st[key];
                    if (!idres.IsErased)
                        return idres;
                    foreach (ObjectId id in st)
                    {
                        if (!id.IsErased)
                        {
                            SymbolTableRecord str = (SymbolTableRecord)trans.GetObject(id, OpenMode.ForRead);
                            if (str.Name == key)
                                return id;
                        }
                    }
                }
            }
            return ObjectId.Null;
        }

        internal static Point3dCollection GetRecPc(Point3d startPt, Point3d endPt)
        {
            Point3dCollection pc = new Point3dCollection();
            pc.Add(startPt);
            pc.Add(new Point3d(endPt.X, startPt.Y, 0));
            pc.Add(endPt);
            pc.Add(new Point3d(startPt.X, endPt.Y, 0));
            return pc;
        }

        internal static IList<Point3dCollection> GetRecPcs(string tips)
        {
            PlanMapService ps = new PlanMapService();
            IList<Point3dCollection> list = new List<Point3dCollection>();
            DBText dbText = ps.SelectText(tips);
            Point3d startPt = ps.PickupPoint("拾取存图左上角的一点");
            Point3d endPt = ps.PickupPoint("拾取存图右下角的一点");

            //得到矩形框
            Point3dCollection recPc = GetRecPc(startPt, endPt);
            Point3dCollection pcDi = ps.GetDistanceReduce(dbText.Position, recPc);

            Point3dCollection titlePoints = utils.findEntityPoints(dbText.TextString);

            foreach (Point3d pt in titlePoints)
            {
                Point3dCollection onePc = ps.GetDistanceReduce(pt, pcDi);
                list.Add(onePc);
            }
            return list;

        }
        public static double Angle(Point3d startPoint, Point3d endPoint)
        {
            //两点的x、y值
            double x = endPoint.X - startPoint.X;
            double y = endPoint.Y - startPoint.Y;

            //斜边长度
            double hypotenuse = Math.Sqrt(Math.Pow(x, 2f) + Math.Pow(y, 2f));

            //求出弧度
            double cos = x / hypotenuse;
            double radian = Math.Acos(cos);

            //用弧度算出角度    
            double angle = 180 / (Math.PI / radian);

            if (y < 0)
            {
                angle = -angle;
            }
            else if ((y == 0) && (x < 0))
            {
                angle = 180;
            }
            return angle;
        }

        internal static Line CreateLine(Transaction trans, Point3d startPt, Point3d endPt, string layer)
        {
            Line line = new Line(startPt, endPt);
            line.Layer = layer;
            SaveEntity(trans, line);
            return line;
        }
        public static IList<DBText> SeletTexts()
        {
            TypedValue[] types = new TypedValue[]
            {new TypedValue((int)DxfCode.Start, "TEXT")
               
             };
            IList<DBText> collection = new List<DBText>();
            SelectionFilter sf = new SelectionFilter(types);
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionResult ents = ed.GetSelection(sf);
            DBObjectCollection texts = utils.GetEntityCollection(ents);
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (DBObject ob in texts)
                {
                    DBText text = (DBText)ob;
                    {
                        collection.Add(text);
                    }


                }
                trans.Commit();
            }
            return collection;

        }

        public static Circle CreateCircle(Point3d center, double big)
        {


            Circle circle = new Circle(center, Vector3d.ZAxis, big);
            return circle;

            
        }
        /// <summary>
        /// 检查 pl 是否包含 jmdPL
        /// </summary>
        /// <param name="pl"></param>
        /// <param name="jmdPL"></param>
        /// <returns></returns>
        public static bool CheckPLContains(Polyline pl, Polyline jmdPL)
        {

           
            if (pl == jmdPL)
            {
                return false;
            }
            Autodesk.AutoCAD.DatabaseServices.Region plRegion = CreateRegion(pl);
            Region re = CADUtils.CreateRegion(jmdPL);
            if (re != null && re.Area !=0)
            {
                double area = plRegion.Area;
                if (plRegion.Area == 0)
                {
                    return false;
                }
                try
                {
                    plRegion.BooleanOperation(BooleanOperationType.BoolUnite, re);
                }
                catch
                {
                    return false;
                }
                
                if (plRegion.Area < area + 0.1)
                {
                    return true;

                }

            }
            else
            {
                return true;
            }
            return false;
        }
/// <summary>
/// 点偏移
/// </summary>
/// <param name="pt1"></param>
/// <param name="pt2"></param>
/// <param name="di"></param>
/// <returns></returns>
        public static Point3dCollection PointOffset(Point3d pt1, Point3d pt2,double di)
        {

            Point3dCollection jzxPoints = new Point3dCollection();
            //jzxPoints.Add(pt1);
            //jzxPoints.Add(pt2);
            Polyline acPoly = new Polyline();
            //加入点
            acPoly.AddVertexAt(0, new Point2d(pt1.X, pt1.Y), 0, 0, 0);
            acPoly.AddVertexAt(1, new Point2d(pt2.X, pt2.Y), 0, 0, 0);
            //本多段线向内偏移0.05
            DBObjectCollection collectionN = acPoly.GetOffsetCurves(di);
            foreach (DBObject dbo in collectionN)
            {
                if (dbo is Polyline)
                {
                    Polyline line = (Polyline)dbo;
                    //得到所有点，在放到点集合中
                    Point3dCollection ps = new Point3dCollection();

                    line.GetStretchPoints(ps);
                    for (int d = 0; d < ps.Count; d++)
                    {
                        jzxPoints.Add(ps[d]);
                    }
                }
            }

            //向外偏移
            DBObjectCollection collection = acPoly.GetOffsetCurves(-di);
            foreach (DBObject dbo in collection)
            {
                if (dbo is Polyline)
                {
                    Polyline line = (Polyline)dbo;
                    //得到所有点，在放到点集合中
                    Point3dCollection ps = new Point3dCollection();

                    line.GetStretchPoints(ps);
                    for (int d = ps.Count - 1; d >= 0; d--)
                    {
                        jzxPoints.Add(ps[d]);
                    }
                }
            }
            //点结合生成新多段线

            jzxPoints = ScalePoints(jzxPoints, utils.GetZpoint(pt1, pt2), 0.6);
           
            return jzxPoints;
        }

        //实体缩放
        public static void Scale(Polyline3d poly, Point3d basicPoint, double scaleFactor)
        {
            Matrix3d mt = Matrix3d.Scaling(scaleFactor, basicPoint);
            poly.TransformBy(mt);

        }
        public static Point3dCollection ScalePoints(Point3dCollection points, Point3d basicPoint, double scale)
        {
        
            Point3dCollection pc = new Point3dCollection();
            double x = basicPoint.X;
            double y = basicPoint.Y;
            double xd;
            double yd;
            foreach (Point3d ptSrc in points)
            {
                xd = ptSrc.X - x;
                yd = ptSrc.Y - y;
                pc.Add(new Point3d(xd * scale + x,yd * scale + y,0));

            }

           
            return pc;
        }
        public static DBObjectCollection[] GetScalePoints(Polyline polyline)
        {
            JzdService jzdService = new JzdService();

            DBObjectCollection[] dbcs = new DBObjectCollection[4];
             Point3dCollection point3dCollection = new Point3dCollection();
            polyline.GetStretchPoints(point3dCollection);
            double[] dou = utils.FindFourPoints(point3dCollection);
            Point3dCollection point3dCollection2 = utils.createFourPoint(dou);
            Polyline3d polyline3d = utils.CreatePolye(point3dCollection2, true);
            IList<Zdinfo> list = JzdService.DBCToZdinfo(utils.SelectJzd(point3dCollection2));
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
            
            IList<string[]> list2 = new List<string[]>();
            double num3 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[1]]);
            double num4 = utils.PointAngle(point3dCollection[array[0]], point3dCollection[array[3]]);
            double num5 = num3 - num4;
            for (int i = 0; i < 4; i++)
            {
                int start = array[i];
                int end = array[i + 1];
                DBObjectCollection dbc = jzdService.SelectJzd(polyline, point3dCollection, array, i, 0.02, 0.9, 1);
                dbc.Remove(polyline);
                IList<Zdinfo> list3 = JzdService.DBCToZdinfo(dbc);
             
               ;
                if (num5 < 0.0 && num3 < 40.0)
                {
                    switch (i)
                    {
                        case 0:
                            dbcs[0] = dbc;
                            break;
                        case 1:
                            dbcs[1] = dbc;
                            break;
                        case 2:
                            dbcs[2] = dbc;
                            break;
                        case 3:
                            dbcs[3] = dbc;
                            break;
                    }
                }
                else
                {
                    switch (i)
                    {
                        case 0:
                            dbcs[1] = dbc;
                            break;
                        case 1:
                            dbcs[2] = dbc;
                            break;
                        case 2:
                            dbcs[3] = dbc;
                            break;
                        case 3:
                            dbcs[0] = dbc;
                            break;
                    }
                }
            }
            return dbcs;
        }


        internal static void ZhengShiZhongDiTu()
        {
            throw new Exception("The method or operation is not implemented.");
        }



        internal static Point3dCollection CreatePLByOnePoint(Point3d pt,double di)
        {
            Point3dCollection pc = new Point3dCollection();
            double x = pt.X;
            double y = pt.Y;
            pc.Add( new Point3d(x - di, y + di, 0));
            pc.Add(new Point3d(x + di, y + di, 0));
            pc.Add(new Point3d(x + di, y - di, 0));
            pc.Add(new Point3d(x - di, y - di, 0));
          
            return pc;
        }
    }
}
