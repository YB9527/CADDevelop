using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using YanBo_CG;

namespace YanBo
{
    internal class PlanMapService
    {
        private static Utils utils = new Utils();

        public Point3dCollection GetDrawingFramePoints(string tips)
        {
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            Utils utils = new Utils();
            Point3dCollection point3dCollection = utils.findEntityPoints(tips);
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "LINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            Point3dCollection point3dCollection2 = new Point3dCollection();
            using (utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (Point3d point3d in point3dCollection)
                {
                    Point3d point3d2 = new Point3d(point3d.X + 1.0, point3d.Y, 0.0);
                    Point3d point3d3 = new Point3d(point3d.X + 1.0, point3d.Y - 5.0, 0.0);
                    Point3d point3d4 = new Point3d(point3d.X, point3d.Y - 5.0, 0.0);
                    Point3dCollection point3dCollection3 = new Point3dCollection();
                    point3dCollection3.Add(point3d);
                    point3dCollection3.Add(point3d2);
                    point3dCollection3.Add(point3d3);
                    point3dCollection3.Add(point3d4);
                    Polyline3d polyline3d = utils.CreatePolye(point3dCollection3, true);
                    Point3dCollection point3dCollection4 = new Point3dCollection();
                    polyline3d.GetStretchPoints(point3dCollection4);
                    PromptSelectionResult ents = editor.SelectCrossingPolygon(point3dCollection4, selectionFilter);
                    DBObjectCollection entityCollection = utils.GetEntityCollection(ents);
                    foreach (Entity entity in entityCollection)
                    {
                        if (entity is Line)
                        {
                            Line line = entity as Line;
                            if (line.StartPoint.X > line.EndPoint.X)
                            {
                                point3dCollection2.Add(line.EndPoint);
                            }
                            else
                            {
                                point3dCollection2.Add(line.StartPoint);
                            }
                        }
                    }
                }
            }
            return point3dCollection2;
        }

        public DBText SelectText(string tips)
        {
            DBText textEntity = utils.getTextEntity(tips);
            if (textEntity == null || textEntity.TextString.Equals(""))
            {
                throw new Exception("你选择的不是单个的文字");
            }
            return textEntity;
        }

        public List<Nf> CheckPlanArea()
        {
            Editor editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            List<Nf> list = new List<Nf>();
            string textString = this.SelectText("请选择房产平面图单个图框标题文字").TextString;
            Point3dCollection drawingFramePoints = this.GetDrawingFramePoints(textString);
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (Point3d point3d in drawingFramePoints)
                {
                    Nf nf = new Nf();
                    double num = 0.0;
                    Point3d pt = new Point3d(point3d.X + 4.54, point3d.Y, 0.0);
                    Point3d pt2 = new Point3d(point3d.X + 12.87, point3d.Y - 2.86, 0.0);
                    string windowText = utils.getWindowText(pt, pt2);
                    nf.setZdnum(windowText);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage(string.Concat(new object[]
						{
							"点是：（",
							point3d.X + 4.54,
							",",
							point3d.Y,
							"）  编码没有填写 "
						}));
                    }
                    pt = new Point3d(point3d.X + 4.54, point3d.Y - 2.86, 0.0);
                    pt2 = new Point3d(point3d.X + 12.87, point3d.Y - 5.72, 0.0);
                    windowText = utils.getWindowText(pt, pt2);
                    nf.setCbfmc(windowText);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage("编码是：" + nf.getZdnum() + ", 名字没有填写 ");
                    }
                    double x = point3d.X;
                    double y = point3d.Y;
                    Point3dCollection point3dCollection = new Point3dCollection();
                    point3dCollection.Add(new Point3d(x + 0.71, y - 6.72, 0.0));
                    point3dCollection.Add(new Point3d(x + 13.43, y - 5.96, 0.0));
                    point3dCollection.Add(new Point3d(x + 13.43, y - 0.23, 0.0));
                    point3dCollection.Add(new Point3d(x + 63.88, y - 0.23, 0.0));
                    point3dCollection.Add(new Point3d(x + 63.88, y - 59.18, 0.0));
                    point3dCollection.Add(new Point3d(x + 52.75, y - 59.18, 0.0));
                    point3dCollection.Add(new Point3d(x + 52.75, y - 65.9, 0.0));
                    point3dCollection.Add(new Point3d(x + 0.24, y - 65.9, 0.0));
                    Polyline3d poly = utils.CreatePolye(point3dCollection, true);
                    DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(poly, array);
                    foreach (Entity entity in corssEntityCollection)
                    {
                        if (entity is Polyline)
                        {
                            Polyline polyline = entity as Polyline;
                            if (polyline.Closed)
                            {
                                List<string> windowText2 = utils.getWindowText(polyline);
                                string linetype = polyline.Linetype;
                                if (linetype.ToUpper().Equals("CONTINUOUS") || linetype.ToUpper().Equals("BYLAYER"))
                                {
                                    if (!windowText2.Contains("1"))
                                    {
                                        utils.CreateCircle(utils.GetPolyCore(polyline), 1, "查看线型");
                                    }
                                    num += Math.Round(polyline.Area, 2);
                                    double num2 = Math.Round(3.454, 2);
                                    num2 = Math.Round(3.455, 2);
                                    num2 = Math.Round(3.456, 2);
                                }
                                else if (linetype.ToUpper().Equals("X5"))
                                {
                                    num += Math.Round(polyline.Area / 2.0, 2);
                                    if (!windowText2.Contains("2"))
                                    {
                                        utils.CreateCircle(utils.GetPolyCore(polyline), 1, "查看线型");
                                    }
                                }
                            }
                            DBObjectCollection corssEntityCollection2 = utils.GetCorssEntityCollection(polyline, array);
                            foreach (Entity entity2 in corssEntityCollection2)
                            {
                                if (!(entity == entity2))
                                {
                                    if (entity2 is Polyline)
                                    {
                                        Polyline pl = entity2 as Polyline;
                                        Region region = utils.RegionOpreation(polyline, pl);
                                        if (region != null && region.Area > 0.1)
                                        {
                                            editor.WriteMessage("\n宗地编码:" + nf.getCbfmc() + ":面积可能有问题");
                                            utils.CreateCircle(utils.GetPolyCore(polyline), 2, "查看是否重叠");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    nf.setZmj(string.Format("{0:F}", num));
                    list.Add(nf);
                }
                transaction.Commit();
            }
            return list;
        }

        public void createNfZmj(List<Nf> nfs)
        {
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook();
            ISheet sheet = hSSFWorkbook.CreateSheet("Sheet1");
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("承包方编码");
            row.CreateCell(1).SetCellValue("承包方名称");
            row.CreateCell(2).SetCellValue("总面积");
            for (int i = 0; i < nfs.Count; i++)
            {
                Nf nf = nfs[i];
                row = sheet.CreateRow(i + 1);
                for (int j = 0; j < 3; j++)
                {
                    ICell cell = row.CreateCell(j);
                    switch (j)
                    {
                        case 0:
                            cell.SetCellValue(nf.getZdnum());
                            break;
                        case 1:
                            cell.SetCellValue(nf.getCbfmc());
                            break;
                        case 2:
                            cell.SetCellValue(nf.getZmj());
                            break;
                    }
                }
            }
            string text = "d:\\宅基地\\统计平面图面积.xls";
            FileStream fileStream = new FileStream(text, FileMode.Create);
            hSSFWorkbook.Write(fileStream);
            fileStream.Close();
            utils.GetEditor().WriteMessage("\n平面图统计结果生成在：" + text);
        }

        public void SetPolyText(string layerName, Polyline3d poly)
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(8, layerName),
				new TypedValue(0, "LINE")
			};
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(poly, types);
            foreach (Entity entity in corssEntityCollection)
            {
                if (entity is DBText)
                {
                    DBText dBText = entity as DBText;
                }
            }
        }

        public void SplitDrawingForm(string berforStr, string afterStr, string dir)
        {
            Editor editor = utils.GetEditor();
            Document document = utils.GetDocument();
            DBText dBText = this.SelectText("请选择另存图框的标题文字\n");
            Point3d pt = this.PickupPoint("拾取标题下直线左边的一点");
            Point3d point3d = this.PickupPoint("拾取存图左上角的一点");
            Point3d point3d2 = this.PickupPoint("拾取存图右下角的一点");
            Point3dCollection point3dCollection = new Point3dCollection();
            point3dCollection.Add(point3d);
            point3dCollection.Add(point3d2);
            Point3dCollection points = utils.createFourPoint(utils.FindFourPoints(point3dCollection));
            Point3dCollection distanceReduce = this.GetDistanceReduce(pt, points);
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Point3dCollection drawingFramePoints = this.GetDrawingFramePoints(dBText.TextString);
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                foreach (Point3d pt2 in drawingFramePoints)
                {
                    string text = null;
                    Point3dCollection point3dCollection2 = this.GetDistanceReduce(pt2, distanceReduce);
                    point3dCollection2 = utils.CreatePolye(point3dCollection2);
                    DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(point3dCollection2);
                    foreach (DBObject dBObject in corssEntityCollection)
                    {
                        if (dBObject is DBText)
                        {
                            DBText dBText2 = (DBText)dBObject;
                            text = dBText2.TextString;
                            if (text != null && text.Length == 19 && (text.Contains("JC") || text.Contains("JB")))
                            {
                                break;
                            }
                        }
                        text = null;
                    }
                    if (text == null)
                    {
                        Point3d pt3 = new Point3d(pt2.X + 4.54, pt2.Y, 0.0);
                        Point3d pt4 = new Point3d(pt2.X + 12.87, pt2.Y - 2.86, 0.0);
                        text = utils.getWindowText(pt3, pt4);
                        if (text == null || "".Equals(text))
                        {
                            TypedValue[] qSSurface = utils.GetQSSurface(corssEntityCollection);
                            if (qSSurface != null && qSSurface.Length > 4)
                            {
                                text = qSSurface[2].Value.ToString();
                            }
                            else
                            {
                                editor.WriteMessage(string.Concat(new object[]
								{
									"点是：（",
									pt2.X + 4.54,
									",",
									pt2.Y,
									"）  编码没有填写 "
								}));
                            }
                        }
                    }
                    string fileName = string.Concat(new string[]
					{
						dir,
						"/",
						berforStr,
						text,
						afterStr,
						".dwg"
					});
                    utils.WClone(corssEntityCollection, fileName);
                }
                transaction.Commit();
            }
        }

        private Point3dCollection GetDistanceReduce(Point3dCollection fourPoints)
        {
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(fourPoints);
            foreach (DBObject dBObject in corssEntityCollection)
            {
            }
            return null;
        }

        public Point3dCollection GetDistanceAdd(Point3d pt, Point3dCollection points)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            double x = pt.X;
            double y = pt.Y;
            double z = pt.Z;
            foreach (Point3d point3d in points)
            {
                point3dCollection.Add(new Point3d(x + point3d.X, y + point3d.Y, z + point3d.Z));
            }
            return point3dCollection;
        }

        public Point3dCollection GetDistanceReduce(Point3d pt, Point3dCollection points)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            double x = pt.X;
            double y = pt.Y;
            foreach (Point3d point3d in points)
            {
                point3dCollection.Add(new Point3d(x - point3d.X, y - point3d.Y, 0.0));
            }
            return point3dCollection;
        }

        /// <summary>
        ///拾取点
        /// </summary>
        /// <param name="message">提示信息</param>
        /// <returns>点</returns>
        public Point3d PickupPoint(String message)
        {
            Editor ed = utils.GetEditor();
            PromptPointResult pr = ed.GetPoint(message);

            if (pr.Status == PromptStatus.OK)
            {
                return pr.Value;
            }
            else
            {
                throw new System.Exception("你没有选择点\n");
            }

        }
        internal void RegulatePlanMap()
        {
            IList<string> list = new List<string>();
            list.Add("JMD");
            list.Add("彩钢");
            list.Add("彩钢棚");
            list.Add("假三层");
            list.Add("集体土地所有权（清风乡、福田乡、海螺乡、坛罐乡、玉成乡）");
            list.Add("楼梯间");
            list.Add("卫浴");
            list.Add("卫浴室");
            list.Add("0");
            list.Add("JZD");
            list.Add("ZBTZ");
            list.Add("DMTZ");
            list.Add("DLSS");
            list.Add("图块");
            list.Add("KZD");
            list.Add("ZDH");
            list.Add("户主注记");
            list.Add("DLDW");
            Utils.CreateLayer("图块");
            using (Transaction transaction = utils.GetTransaction())
            {
                Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                foreach (string current in zdDic_zdnum.Keys)
                {
                    DBObjectCollection dbc = zdDic_zdnum[current];
                    Polyline polyline = this.FindQSX(current, dbc);
                    if (polyline == null)
                    {
                        throw new Exception("宗地编码：" + current + "，没有这个权属线");
                    }
                    Zdinfo zdinfo = new Zdinfo();
                    zdinfo.setZdnum(current);
                    ResultBuffer xDataForApplication = polyline.GetXDataForApplication("SOUTH");
                    zdinfo.setQuanli(xDataForApplication.AsArray()[3].Value.ToString());
                    zdinfo.Polyline = polyline;
                    this.RegulatePlanMap(transaction, dbc, zdinfo);
                    this.DeleteExPl(transaction, list, polyline, dbc);
                }
                transaction.Commit();
            }
        }

        private Polyline FindQSX(string zdnum, DBObjectCollection dbc)
        {
            Polyline result;
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is Polyline)
                {
                    ResultBuffer xDataForApplication = dBObject.GetXDataForApplication("SOUTH");
                    if (!(xDataForApplication == null))
                    {
                        TypedValue[] array = xDataForApplication.AsArray();
                        if (array.Length > 2)
                        {
                            if (array[2].Value.ToString().Equals(zdnum))
                            {
                                result = (Polyline)dBObject;
                                return result;
                            }
                        }
                    }
                }
            }
            result = null;
            return result;
        }
        /// <summary>
        ///删除图框外不要的东西
        /// </summary>
        /// <param name="trans"></param>
        /// <param name="deleteLayer"></param>
        /// <param name="qsx"></param>
        /// <param name="dbc"></param>
        private void DeleteExPl(Transaction trans, IList<string> deleteLayer, Polyline qsx, DBObjectCollection dbc)
        {
            DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(qsx);
            IList<ObjectId> dBObjectCollectionIds = this.GetDBObjectCollectionIds(corssEntityCollection);
            string[] deleteText = new string[]{
                "砖","木","破","混","土","阳台","坝（土）","1","2","3","4","5","6","072"
            };
            foreach (DBObject dBObject in dbc)
            {
                bool flag = false;
                if (!dBObjectCollectionIds.Contains(dBObject.Id))
                {
                    
                    if (dBObject is Polyline)
                    {
                        Polyline polyline = trans.GetObject(dBObject.ObjectId, OpenMode.ForWrite) as Polyline;
                        if (deleteLayer.Contains(polyline.Layer))
                        {
                            polyline.Erase();
                           // flag = true;
                        }
                    }
                    else if (dBObject is BlockReference)
                    {

                        BlockReference blockReference = (BlockReference)dBObject;
                        if (deleteLayer.Contains(blockReference.Layer))
                        {
                            blockReference.Erase();
                            flag = true;
                        }

                    }
                    else if (dBObject is Line)
                    {
                        Line line = (Line)dBObject;
                        if (deleteLayer.Contains(line.Layer))
                        {
                            line.Erase();
                            flag = true;
                        }
                    }
                    else if (dBObject is DBText)
                    {
                        //删除不要的文字
                        string text = ((DBText)dBObject).TextString;
                        if (Array.IndexOf(deleteText, text) != -1)
                        {
                            if (text == "2" && ((DBText)dBObject).Height < 0.62 )
                            {//|| ((DBText)dBObject).Layer == "JZD"
                                continue;
                               
                            }
                            dBObject.Erase();
                            flag = true;
                        }
                    }
                }
                else
                {
                    if (dBObject is DBText)
                    {
                        string text = ((DBText)dBObject).TextString;
                        ((DBText)dBObject).Height = 1.2;
                        double t;
                        if(text.Contains(".") && double.TryParse(text,out t))
                        {
                            dBObject.Erase();
                            flag = true;
                        }
                    }
                }
                if (flag)
                {
                    dbc.Remove(dBObject);
                }
            }
        }

        private IList<ObjectId> GetDBObjectCollectionIds(DBObjectCollection dbc)
        {
            IList<ObjectId> list = new List<ObjectId>();
            foreach (DBObject dBObject in dbc)
            {
                list.Add(dBObject.ObjectId);
            }
            return list;
        }

        private void RegulatePlanMap(Transaction trans, DBObjectCollection dbc, Zdinfo zdinfo)
        {
            IList<string> list = new List<string>();
            list.Add("砖");
            list.Add("砼");
            list.Add("简");
            list.Add("水泥");
            list.Add("土");
            list.Add("2");
            list.Add("3");
            list.Add("4");
            list.Add("5");
            Polyline polyline = zdinfo.Polyline;
            polyline.ConstantWidth = (0.15);
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is DBText)
                {
                    DBText dBText = trans.GetObject(dBObject.ObjectId, OpenMode.ForWrite) as DBText;
                    double height = dBText.Height;
                    double widthFactor = dBText.WidthFactor;
                    string textString = dBText.TextString;
                    TextStyleTableRecord textStyleTableRecord = trans.GetObject(dBText.TextStyle, 0) as TextStyleTableRecord;
                    string name = textStyleTableRecord.Name;
                    TextHorizontalMode horizontalMode = dBText.HorizontalMode;
                    TextVerticalMode verticalMode = dBText.VerticalMode;
                    double num;
                    if (dBText.Layer.Equals("权利人"))
                    {
                        dBText.Erase();
                    }
                    else if (double.TryParse(textString, out num) && num <= 0.04)
                    {
                        dBText.Erase();
                    }
                    else if (list.Contains(textString))
                    {
                        CADUtils.SetText(dBText, textString, "宋体", 1.5, 0.8);
                    }
                    else if (height == 1.75 && widthFactor == 0.8)
                    {
                        dBText.Height = (1.75);
                        dBText.WidthFactor = (1.0);
                        if (textString.Contains("户"))
                        {
                            Point3d position = dBText.Position;
                            Vector3d vectorTo = position.GetVectorTo(new Point3d(position.X + 1.18, position.Y + 3.19, 0.0));
                            IntegerCollection integerCollection = new IntegerCollection(0);
                            dBText.MoveStretchPointsAt(integerCollection, vectorTo);
                        }
                    }
                    else if (height == 1.5 && widthFactor == 0.8)
                    {
                        if (textString.StartsWith("JC") || textString.StartsWith("JB"))
                        {
                            dBText.Height = (1.75);
                            dBText.WidthFactor = (1.0);
                            dBText.TextString = (zdinfo.getQuanli());
                        }
                        else if (textString.Equals("072"))
                        {
                            dBText.Height = (1.75);
                            dBText.WidthFactor = (1.0);
                            dBText.TextString = (zdinfo.getZdnum());
                        }
                    }
                    else if (height == 1.0 && widthFactor == 0.8)
                    {
                        if (verticalMode == TextVerticalMode.TextVerticalMid)
                        {
                            dBText.Erase();
                        }
                        else if (horizontalMode != TextHorizontalMode.TextMid)
                        {
                            dBText.Height = (1.5);
                            dBText.WidthFactor = (1.0);
                        }
                    }
                }
                else if (dBObject is Line)
                {
                    Line line = trans.GetObject(dBObject.ObjectId, OpenMode.ForWrite) as Line;
                    Point3d startPoint = line.StartPoint;
                    Point3d endPoint = line.EndPoint;
                    if (line.EndParam == 10.5 && CADUtils.Angle(startPoint, endPoint) == 180.0)
                    {
                        line.Erase();
                    }
                    if (line.EndParam == 2.0 && CADUtils.Angle(startPoint, endPoint) == 0.0)
                    {
                        CADUtils.CreateLine(trans, new Point3d(startPoint.X - 2.5, startPoint.Y, 0.0), new Point3d(endPoint.X + 2.5, endPoint.Y, 0.0), "TK");
                        line.Erase();
                    }
                }
                else if (dBObject is BlockReference)
                {
                    BlockReference blockReference = trans.GetObject(dBObject.ObjectId, OpenMode.ForWrite) as BlockReference;
                    blockReference.Layer = ("图块");
                }
            }
        }

        internal Dictionary<string, DBObjectCollection> GetZdDic_zdnum19()
        {
            Dictionary<string, DBObjectCollection> dictionary = new Dictionary<string, DBObjectCollection>();
            IList<Point3dCollection> recPcs = CADUtils.GetRecPcs("请选择图框的标题文字\n");
            foreach (Point3dCollection current in recPcs)
            {
                int count = dictionary.Count;
                DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(current);
                foreach (DBObject dBObject in corssEntityCollection)
                {
                    if (dBObject is DBText)
                    {
                        DBText dBText = (DBText)dBObject;
                        string textString = dBText.TextString;
                        if (textString != null && textString.Length == 19 && (textString.Contains("JC") || textString.Contains("JB")))
                        {
                            if (dictionary.ContainsKey(textString))
                            {
                                MessageBox.Show("宗地编码重复：" + textString);
                            }
                            else
                            {
                                dictionary.Add(textString, corssEntityCollection);
                            }
                            break;
                        }
                    }
                }
                if (count == dictionary.Count)
                {
                    throw new Exception("有图没有19位的编码");
                }
            }
            return dictionary;
        }

        internal void FlushPlanMap()
        {
            string text = new FormService().OpenFile("选择建设用地表", "Excel(*.xls)|*.xlsx;*.xls");
            if (!text.Equals(""))
            {
                Dictionary<string, Jsyd> dictionary = JsydService.ExcelToJsyd(text);
                using (Transaction transaction = utils.GetTransaction())
                {
                    Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                    foreach (string current in zdDic_zdnum.Keys)
                    {
                        Jsyd jsyd;
                        if (dictionary.TryGetValue(current, out jsyd))
                        {
                            this.FlushPlanMap(transaction, zdDic_zdnum[current], jsyd);
                        }
                        else
                        {
                            MessageBox.Show("建设用地中没有找到此宗地编码，无法更新：" + current);
                        }
                    }
                    transaction.Commit();
                }
            }
        }

        private void FlushPlanMap(Transaction trans, DBObjectCollection dbc, Jsyd jsyd)
        {
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is DBText)
                {
                    DBText dBText = trans.GetObject(dBObject.ObjectId, OpenMode.ForWrite) as DBText;
                    double height = dBText.Height;
                    double widthFactor = dBText.WidthFactor;
                    string textString = dBText.TextString;
                    TextStyleTableRecord textStyleTableRecord = trans.GetObject(dBText.TextStyle, 0) as TextStyleTableRecord;
                    string name = textStyleTableRecord.Name;
                    TextHorizontalMode horizontalMode = dBText.HorizontalMode;
                    TextVerticalMode verticalMode = dBText.VerticalMode;
                    double num;
                    if (height == 1.75 && widthFactor == 1.0 && double.TryParse(textString, out num))
                    {
                        dBText.TextString = (jsyd.Zdmj.ToString("f2"));
                    }
                }
            }
            IList<ObjectId> list = this.FindDbText("土地权利人", FileSelectRelation.StartsWith, dbc);
            DBText dBText2 = trans.GetObject(list[0], OpenMode.ForWrite) as DBText;
            Point3d position = dBText2.Position;
            double syqmj = jsyd.Syqmj;
            IList<ObjectId> list2 = this.FindDbText("使用权面积中", FileSelectRelation.StartsWith, dbc);
            if (list2.Count > 1)
            {
                for (int i = 1; i < list2.Count; i++)
                {
                    trans.GetObject(list2[i], OpenMode.ForWrite).Erase();
                }
            }
            if (syqmj != 0.0)
            {
                DBText dBText;
                if (list2.Count == 0)
                {
                    dBText = (dBText2.Clone() as DBText);
                    Vector3d vectorTo = position.GetVectorTo(new Point3d(position.X - 38.54, position.Y - 103.95, 0.0));
                    IntegerCollection integerCollection = new IntegerCollection(0);
                    dBText.MoveStretchPointsAt(integerCollection, vectorTo);
                    dBText.Height = (1.5);
                    CADUtils.SaveEntity(trans, dBText);
                }
                else
                {
                    dBText = (trans.GetObject(list2[0], OpenMode.ForWrite) as DBText);
                }
                dBText.TextString = ("使用权面积中批准拨用宅基地面积" + syqmj.ToString("f2") + "平方米");
            }
            else if (list2.Count == 1)
            {
                trans.GetObject(list2[0], OpenMode.ForWrite).Erase();
            }
            double czmj = jsyd.Czmj;
            IList<ObjectId> list3 = this.FindDbText("其他集体建设用地", FileSelectRelation.StartsWith, dbc);
            if (list3.Count > 1)
            {
                for (int i = 1; i < list2.Count; i++)
                {
                    trans.GetObject(list3[i], OpenMode.ForWrite).Erase();
                }
            }
            if (czmj != 0.0)
            {
                DBText dBText;
                if (list3.Count == 0)
                {
                    dBText = (dBText2.Clone() as DBText);
                    Vector3d vectorTo = position.GetVectorTo(new Point3d(position.X - 38.54, position.Y - 107.16, 0.0));
                    IntegerCollection integerCollection = new IntegerCollection(0);
                    dBText.MoveStretchPointsAt(integerCollection, vectorTo);
                    dBText.Height = (1.5);
                    CADUtils.SaveEntity(trans, dBText);
                }
                else
                {
                    dBText = (trans.GetObject(list3[0], OpenMode.ForWrite) as DBText);
                }
                dBText.TextString = ("其他集体建设用地面积" + czmj.ToString("f2") + "平方米");
            }
            else if (list3.Count > 0)
            {
                for (int i = 0; i < list3.Count; i++)
                {
                    trans.GetObject(list3[i], OpenMode.ForWrite).Erase();
                }
            }
        }

        private IList<ObjectId> FindDbText(string str, FileSelectRelation relation, DBObjectCollection dbc)
        {
            IList<ObjectId> list = new List<ObjectId>();
            foreach (DBObject dBObject in dbc)
            {
                if (dBObject is DBText)
                {
                    DBText dBText = dBObject as DBText;
                    string textString = dBText.TextString;
                    switch (relation)
                    {
                        case FileSelectRelation.StartsWith:
                            if (textString.StartsWith(str))
                            {
                                list.Add(dBObject.ObjectId);
                            }
                            break;
                        case FileSelectRelation.Contains:
                            if (textString.Contains(str))
                            {
                                list.Add(dBObject.ObjectId);
                            }
                            break;
                        case FileSelectRelation.EndsWith:
                            if (textString.EndsWith(str))
                            {
                                list.Add(dBObject.ObjectId);
                            }
                            break;
                        case FileSelectRelation.Equals:
                            if (textString.Equals(str))
                            {
                                list.Add(dBObject.ObjectId);
                            }
                            break;
                    }
                }
            }
            return list;
        }

        internal void DeleteZdnum2()
        {
            using (Transaction transaction = utils.GetTransaction())
            {
                Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                foreach (string current in zdDic_zdnum.Keys)
                {
                    DBObjectCollection dbc = zdDic_zdnum[current];
                    IList<ObjectId> list = this.FindDbText(current, FileSelectRelation.Equals, dbc);
                    if (list.Count == 2)
                    {
                        DBText dBText = transaction.GetObject(list[0], OpenMode.ForWrite) as DBText;
                        DBText dBText2 = transaction.GetObject(list[1], OpenMode.ForWrite) as DBText;
                        if (dBText.Position.Y > dBText2.Position.Y)
                        {
                            dBText2.Erase();
                        }
                        else
                        {
                            dBText.Erase();
                        }
                    }
                }
                transaction.Commit();
            }
        }

        internal void CopyZDT()
        {
            using (utils.GetTransaction())
            {
                Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                Point3d point3d = this.PickupPoint("选择复制的基点");
                Point3d point3d2 = this.PickupPoint("选择复制的目的点");
                Dictionary<string, DBObjectCollection> zdDic_zdnum2 = this.GetZdDic_zdnum19();
                foreach (string current in zdDic_zdnum.Keys)
                {
                }
            }
        }

        /// <summary>
        /// 检建筑物包含
        /// </summary>
        public void CheckJZWContains()
        {
            TypedValue value1 = new TypedValue((int)DxfCode.LayerName, "JMD");
            //1、找到所有建筑物(多段线)
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE"),
                value1
			};
            Point3dCollection errorPC = new Point3dCollection();
            DBObjectCollection dBObjectCollection = utils.SelectDBObjectCollection("选择要检查的多段线", types);
            string[] weiQiang = new string[]
            {
                "144301","144301-1","144301-2","143800","143701"
            };
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in dBObjectCollection)
                {

                    ResultBuffer xDataForApplication = dBObject.GetXDataForApplication("SOUTH");
                    if (xDataForApplication != null)
                    {
                        TypedValue[] array = xDataForApplication.AsArray();
                        if (array.Length <= 3)
                        {

                            if (Array.IndexOf(weiQiang, array[1].Value.ToString()) != -1)
                            {
                                continue;
                            }
                        }
                    }
                    Polyline jmdPL = dBObject as Polyline;
                    //2、所有建筑线 查找包含的多段线

                    DBObjectCollection rs = GetCorssEntityCollection(utils.Scale(jmdPL, 1.01), types);
                    foreach (DBObject r in rs)
                    {
                        //3、做面积 确定是包含的 在中心画圆
                        if (r.ObjectId != jmdPL.ObjectId)
                        {

                            //围墙除外
                            xDataForApplication = r.GetXDataForApplication("SOUTH");
                            if (xDataForApplication != null)
                            {
                                TypedValue[] array = xDataForApplication.AsArray();
                                if (array.Length <= 3)
                                {

                                    if (Array.IndexOf(weiQiang, array[1].Value.ToString()) == -1)
                                    {
                                        errorPC.Add(utils.GetPolyCore(dBObject as Polyline));
                                        break;
                                    }
                                }
                            }



                        }

                    }

                }
            }
            utils.CreateCircle(errorPC, 1, "建筑物包含");
        }
        public static DBObjectCollection GetCorssEntityCollection(Polyline3d poly, TypedValue[] types)
        {
            SelectionFilter selectionFilter = new SelectionFilter(types);
            Point3dCollection point3dCollection = new Point3dCollection();
            poly.GetStretchPoints(point3dCollection);
            PromptSelectionResult ents = utils.GetEditor().SelectWindowPolygon(point3dCollection, selectionFilter);
            return GetEntityCollection(ents);
        }
        public static DBObjectCollection GetEntityCollection(PromptSelectionResult ents)
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

        /// <summary>
        /// 检查宗地图、三个面积、图幅号
        /// </summary>
        internal void CheckZDT()
        {
            Editor editor = utils.GetEditor();
            Document document = utils.GetDocument();
            DBText dBText = this.SelectText("请选择图框的标题文字\n");
            Point3d pt = this.PickupPoint("拾取标题下直线左边的一点");
            Point3d point3d = this.PickupPoint("拾取存图左上角的一点");
            Point3d point3d2 = this.PickupPoint("拾取存图右下角的一点");
            Point3dCollection point3dCollection = new Point3dCollection();
            point3dCollection.Add(point3d);
            point3dCollection.Add(point3d2);
            Point3dCollection points = utils.createFourPoint(utils.FindFourPoints(point3dCollection));


            Point3dCollection distanceReduce = this.GetDistanceReduce(pt, points);
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Point3dCollection drawingFramePoints = this.GetDrawingFramePoints(dBText.TextString);
            if (drawingFramePoints == null || drawingFramePoints.Count == 0)
            {
                return;
            }
            string path = FileUtils.SelectSingleExcelFile("选择建设用地表");
            IList<Jsyd> jsyds = JSYDCustom.ReadExcel(path);
            if (jsyds.Count == 0)
            {
                return;
            }
            Dictionary<string, Jsyd> jsydDic = JSYDCustom.ListToDic(jsyds);
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                foreach (Point3d pt2 in drawingFramePoints)
                {
                    string text = null;
                    Point3dCollection point3dCollection2 = this.GetDistanceReduce(pt2, distanceReduce);
                    point3dCollection2 = utils.CreatePolye(point3dCollection2);
                    DBObjectCollection corssEntityCollection = utils.GetCorssEntityCollection(point3dCollection2);
                    foreach (DBObject dBObject in corssEntityCollection)
                    {
                        if (dBObject is DBText)
                        {
                            DBText dBText2 = (DBText)dBObject;
                            text = dBText2.TextString;
                            if (text != null && text.Length == 19 && (text.Contains("JC") || text.Contains("JB")))
                            {
                                Jsyd jsyd;
                                if (jsydDic.TryGetValue(text, out jsyd))
                                {
                                    JSYDCustom.CheckZDT(jsyd, corssEntityCollection);
                                }
                                else
                                {
                                    editor.WriteMessage("宗地编码在表上没有找到：" + text + "\r\n");
                                }
                            }
                        }

                    }

                }
                transaction.Commit();
            }
        }

        internal void RegulatePlanMap2()
        {

            using (Transaction transaction = utils.GetTransaction())
            {
                Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                foreach (string current in zdDic_zdnum.Keys)
                {
                    DBObjectCollection dbc = zdDic_zdnum[current];
                    Polyline polyline = this.FindQSX(current, dbc);
                    if (polyline == null)
                    {
                        throw new Exception("宗地编码：" + current + "，没有这个权属线");
                    }
                    Zdinfo zdinfo = new Zdinfo();
                    zdinfo.setZdnum(current);
                    ResultBuffer xDataForApplication = polyline.GetXDataForApplication("SOUTH");
                    zdinfo.setQuanli(xDataForApplication.AsArray()[3].Value.ToString());
                    zdinfo.Polyline = polyline;
                    this.RegulatePlanMap(transaction, dbc, zdinfo);

                }
                transaction.Commit();
            }
        }

        /// <summary>
        /// 宗地图文字管理
        /// </summary>
        internal void ZDTTextManger()
        {
            JzdService jzdService = new JzdService();
            IList<string> list = new List<string>();
            list.Add("JMD");
            list.Add("彩钢");
            list.Add("彩钢棚");
            list.Add("假三层");
            list.Add("集体土地所有权（清风乡、福田乡、海螺乡、坛罐乡、玉成乡）");
            list.Add("楼梯间");
            list.Add("卫浴");
            list.Add("卫浴室");
            list.Add("0");
            list.Add("JZD");
            //list.Add("ZBTZ");
            //list.Add("DMTZ");
            //list.Add("DLSS");
            list.Add("图块");
            list.Add("KZD");
            list.Add("ZDH");
            list.Add("户主注记");
            list.Add("DLDW");
            list.Add("坝");
            Utils.CreateLayer("图块");
            //界址点号处理
            //界址点距离处理
            //其他文字处理
            using (Transaction transaction = utils.GetTransaction())
            {
                ObjectId textStyle = utils.CreateSytle1("细等线体", 1.4);
                Dictionary<string, DBObjectCollection> zdDic_zdnum = this.GetZdDic_zdnum19();
                foreach (string current in zdDic_zdnum.Keys)
                {
                    Point3d companyBasicPoint ;
                    DBObjectCollection dbc = zdDic_zdnum[current];
                    Polyline polyline = this.FindQSX(current, dbc);
                    if (polyline == null)
                    {
                        throw new Exception("宗地编码：" + current + "，没有这个权属线");
                    }
                    //DBObjectCollection qsxNei = utils.sele
                    Zdinfo zdinfo = JzdService.ToZdinfo(polyline);
                    zdinfo.setZdnum(current);
                    ResultBuffer xDataForApplication = polyline.GetXDataForApplication("SOUTH");
                    zdinfo.setQuanli(xDataForApplication.AsArray()[3].Value.ToString());
                    zdinfo.Polyline = polyline;
                    //删除多余的线
                    IList<Zdinfo> zdinfos = new List<Zdinfo>();
                     IList<string> deleteQuanLi = new List<string>();
                    this.DeleteExPl(transaction, list, polyline, dbc);

                    foreach (DBObject db in dbc)
                    {
                       
                        //删除邻宗的权属线
                        if (db is Polyline)
                        {
                            Polyline pl = (Polyline)db;
                            if (pl.Layer == "JZD" && !pl.Closed)
                            {
                                //pl.Erase();
                                Zdinfo zdinfoTem = JzdService.ToZdinfo(db);
                                deleteQuanLi.Add(zdinfoTem.getQuanli());
                                deleteQuanLi.Add(zdinfoTem.getZdnum().Substring(6,13));
                                
                                zdinfos.Add(zdinfoTem);
                            }
                        }

                    }
                    double scale=0;
                    foreach (DBObject db in dbc)
                    {
                        if (db is DBText)
                        {
                            DBText dbText = (DBText)db;
                            string text = dbText.TextString;
                            if (text.Equals("    年  月解析法测绘界址点"))
                            {
                                dbText.TextString = "2017年7月解析法测绘界址点";
                                companyBasicPoint = new Point3d(dbText.Position.X,dbText.Position.Y,0);
                                Point3d start = new Point3d(companyBasicPoint.X + 1.72, companyBasicPoint.Y + 31.92, 0);
                                string[] companyTextArray = new string[]{
                                   
                                };
                                for (int a = 0; a < companyTextArray.Length; a++)
                                {
                                    start = new Point3d(start.X, start.Y - 1.72 , 0);
                                    DBText dBText = utils.CreateDBText(companyTextArray[a], start, 0.888, textStyle);
                                    dBText.Layer = ("TK");
                                    //dBText.ColorIndex = 1;
                                    dBText.HorizontalMode = TextHorizontalMode.TextMid;
                                    dBText.AlignmentPoint = dBText.Position;
                                    utils.AddDBText(dBText);
                                }

                            }
                            else if (text.StartsWith("J"))
                            {
                                //删除界址点，自己生成
                                dbText.Erase();
                            }
                            else if (dbText.Layer.Equals("JZD") || dbText.Layer.Equals("权利人") || text.Equals("072"))
                            {
                                //删除界址点距离 ,072
                                dbText.Erase();
                            }
                            else if (text.Equals(zdinfo.getQuanli()) && dbText.Height != 1.25)
                            {
                                 //删除本宗地的权利人 不要
                                dbText.Erase();
                            }
                            else if (text.Contains("JC") && text.Length != 19)
                            {
                                dbText.Erase();
                            }
                            else if (deleteQuanLi.Contains(text))
                            {
                                dbText.Erase();
                            }
                            else
                            {
                                //删除面积
                                double area;
                                if (dbText.Height == 0.8 && dbText.Layer == "TK" && double.TryParse(text,out area))
                                {
                                    dbText.Erase();
                                }
                                else if (text.StartsWith("1:") && text.Length >2 && scale == 0)
                                {
                                    string text1 = text.Split(':')[1];
                                    if (double.TryParse(text1, out scale) )
                                    {
                                        scale = scale / 500;
                                    }
                                    else
                                    {
                                        MessageBox.Show("填写的比例有问题：" + text1);
                                    }
                                }
                            }
                           

                            //删除


                        }
                        else if (db is Line)
                        {
                            Line line = transaction.GetObject(db.ObjectId, OpenMode.ForWrite) as Line;
                            Point3d startPoint = line.StartPoint;
                            Point3d endPoint = line.EndPoint;

                            if ((line.EndParam >= 9.0 && line.EndParam <= 25.1) && (CADUtils.Angle(startPoint, endPoint) == 0.0 || CADUtils.Angle(startPoint, endPoint) == 180))
                            {
                                double x = (startPoint.X + endPoint.X) / 2;
                                //CADUtils.CreateLine(transaction, new Point3d(x - 3.185, startPoint.Y, 0.0), new Point3d(x + 3.185, endPoint.Y, 0.0), "TK");
                                line.Erase();
                            }
                        }

                    }
                    zdinfos.Add(zdinfo);
                 
                    for (int a =0; a < zdinfos.Count;a++)
                    {
                        Zdinfo zd = zdinfos[a];
                        //创建文字
                        if (a != zdinfos.Count - 1)
                        {
                            Polyline pl2 = zd.Polyline;
                            if (pl2.IsErased)
                            {
                                continue;
                            }
                            //zd.Polyline.Erase();
                        }

                        string zdnum = zd.getZdnum();
                        if (zdnum.Length > 7)
                        {
                            zdnum = zdnum.Substring(zdnum.Length - 7, 7);
                        }
                        string qlr = zd.getQuanli();
                        Point3d core = utils.GetPolyCore(zd.Polyline);
                        {
                            if (zdinfo.getBz().IndexOf(qlr) != -1)
                            {
                                core = new Point3d(core.X,core.Y+2,0);
                            }
                            else if (zdinfo.getDz().IndexOf(qlr) != -1)
                            {
                                core = new Point3d(core.X+2, core.Y , 0);
                            }
                            else if (zdinfo.getNz().IndexOf(qlr) != -1)
                            {
                                core = new Point3d(core.X , core.Y-4, 0);
                            }
                            else if (zdinfo.getXz().IndexOf(qlr) != -1)
                            {
                                core = new Point3d(core.X-2, core.Y , 0);
                            }


                        }

                        double x = core.X;
                        CADUtils.CreateLine(transaction, new Point3d(x - 3.185, core.Y, 0.0), new Point3d(x + 3.185, core.Y, 0.0), "TK");


                        core = new Point3d(core.X - 2.14, core.Y + 1.0, 0);
                        DBText dBText = utils.CreateDBText(qlr, core, 1.4, textStyle);
                        dBText.Layer = ("TK");
                        dBText.HorizontalMode = TextHorizontalMode.TextMid;
                        dBText.VerticalMode = TextVerticalMode.TextBottom;
                        dBText.AlignmentPoint = dBText.Position;
                        utils.AddDBText(dBText);



                        core = new Point3d(core.X - 1, core.Y - 3, 0);
                        dBText = utils.CreateDBText(zdnum, core, 1.4, textStyle);
                        dBText.Layer = ("TK");
                        dBText.HorizontalMode = TextHorizontalMode.TextMid;
                        dBText.VerticalMode = TextVerticalMode.TextTop;
                        dBText.AlignmentPoint = dBText.Position;
                        utils.AddDBText(dBText);
                        //最后删除此线 只要不是最后一个（本宗地） 都删除


                    }
                    //生成界址点号
                    Point3dCollection jzds = new Point3dCollection();
                    zdinfo.Polyline.GetStretchPoints(jzds);
                    Point3dCollection jzdFourExPoints = new JzdService().GetJzdFourExPoints(zdinfo.Polyline);
                    jzds.Add(jzds[0]);
                    for (int a = 0; a < jzds.Count-1; a++)
                    {
                        Point3d pt1 = jzds[a];
                        Point3d pt2 = jzds[a + 1];
                        double angle = CADUtils.Angle(pt1, pt2);
                        Double di = utils.CalculateTwoPt(pt1,pt2);
                        Point3d core;
                        Point3d jzxBZ = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                        if (angle >= -20 && angle < 20)
                        {
                            core = new Point3d(pt1.X, pt1.Y + 1.6, 0);
                            if (di < 3)
                            {
                                jzxBZ = new Point3d(jzxBZ.X, jzxBZ.Y + 1.6+1.5, 0); 
                            }
                            else
                            {
                                jzxBZ = new Point3d(jzxBZ.X, jzxBZ.Y + 1.6, 0); 
                            }

                           
                        }
                        else if (angle > -90 && angle<=0)
                        {
                            core = new Point3d(pt1.X+1.2, pt1.Y , 0);
                            jzxBZ = new Point3d(jzxBZ.X+1.2, jzxBZ.Y, 0); 
                        }
                        else if (angle > -180 && angle <= -90)
                        {
                            core = new Point3d(pt1.X , pt1.Y-2, 0);
                            jzxBZ = new Point3d(jzxBZ.X-2 , jzxBZ.Y-3, 0); 
                        }
                        else if (angle >20)
                        {
                            core = new Point3d(pt1.X - 2.6, pt1.Y , 0);
                            if (di < 3)
                            {
                                jzxBZ = new Point3d(jzxBZ.X - 2.6, jzxBZ.Y+0.2 , 0); 
                            }
                            else
                            {
                                jzxBZ = new Point3d(jzxBZ.X - 3.6, jzxBZ.Y+0.2, 0); 
                            }
                           
                        }
                        //点号
                        DBText dBText = utils.CreateDBText("J" + (a + 1), core, 1.0, textStyle);
                        dBText.Layer = ("TK");
                        dBText.ColorIndex = 1;

                        dBText.HorizontalMode = TextHorizontalMode.TextMid;
                        dBText.AlignmentPoint = dBText.Position;

                        utils.AddDBText(dBText);

                        //界址线距离
                        double di2 = utils.CalculateTwoPt(pt2,pt1)*scale;
                        if(di2 <0.5)
                        {
                            continue;
                        }
                        dBText = utils.CreateDBText(di2.ToString("f2"), jzxBZ, 0.8, textStyle);
                        //dBText.SetDatabaseDefaults(utils.GetDabase());


                        dBText.HorizontalMode = TextHorizontalMode.TextMid;
                        dBText.AlignmentPoint = jzxBZ;

                    
                        dBText.Layer = ("TK");
                        dBText.ColorIndex = 1;
                        if (angle < -90 && angle > -180)
                        {
                            angle = angle + 180;
                        }
                        else if (angle > 90)
                        {
                            angle = angle  - 180;
                        }
                        dBText.Rotation =angle/57.3;
                        
                        
                        utils.AddDBText(dBText);
                    }
                   
                }
                transaction.Commit();
            }

        }

        internal void CheckJZWContains2()
        {

            PromptDoubleOptions options = new PromptDoubleOptions("\n请输入整数: ");
            options.DefaultValue = 0.01;
            PromptDoubleResult result = utils.GetEditor().GetDouble(options);
            switch (result.Status)
            {
                case PromptStatus.OK:
                    utils.GetEditor().WriteMessage("\nValue = " + result.Value);
                    break;

                case PromptStatus.Cancel:
                    utils.GetEditor().WriteMessage("\n用户取消了输入");
                    return;

                default:
                    return;
            }
            double inputValue = result.Value;


            TypedValue value1 = new TypedValue((int)DxfCode.LayerName, "JMD");
            //1、找到所有建筑物(多段线)
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE"),
                value1
			};
            Point3dCollection errorPC = new Point3dCollection();
            DBObjectCollection dBObjectCollection = utils.SelectDBObjectCollection("选择要检查的多段线", types);
            string[] weiQiang = new string[]
            {
                "144301","144301-1","144301-2","143800","143701"
            };
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in dBObjectCollection)
                {

                    ResultBuffer xDataForApplication = dBObject.GetXDataForApplication("SOUTH");
                    if (xDataForApplication != null)
                    {
                        TypedValue[] array = xDataForApplication.AsArray();
                        if (array.Length <= 3)
                        {
                            //围墙跳过
                            if (Array.IndexOf(weiQiang, array[1].Value.ToString()) != -1)
                            {
                                continue;
                            }
                        }
                    }
                    Polyline jmdPL = dBObject as Polyline;
                    //2、所有建筑线 查找包含的多段线
                    Point3dCollection pc = new Point3dCollection();
                    jmdPL.GetStretchPoints(pc);
                    for (int a = 0; a < pc.Count; a++)
                    {
                        Point3d pt = pc[a];
                        //用圆去选择
                        //1、找到所有建筑物(多段线)
                        TypedValue[] types1 = new TypedValue[]
			            {
				            new TypedValue(0, "LWPOLYLINE"),
                            new TypedValue((int)DxfCode.LayerName, "JMD")
			            };
                        ///先选择交叉的
                        Point3dCollection selctPC = CADUtils.CreatePLByOnePoint(pt, inputValue);
                        //utils.CreatePolye(selctPC);
                        DBObjectCollection dbcSelect = utils.GetCorssEntityCollection(selctPC);
                        if (dbcSelect.Count > 1)
                        {
                            //循环每一个第每一个面JMD JZD，检查是否有界址点,检查悬挂
                            foreach (DBObject dbo in dbcSelect)
                            {
                                if(dbo == dBObject)
                                {
                                    //如果是自身就 跳过
                                    continue;
                                }
                                if (dbo is Polyline)
                                {
                                    Polyline otherPL = (Polyline)dbo;
                                    if (otherPL.Layer == "JZD" || otherPL.Layer == "JMD")
                                    {
                                        if (!PLContainsPoint(otherPL, pt))
                                        {
                                            //如果不包含就标记错误
                                            errorPC.Add(pt);
                                        }
                                    }
                                }
                            }
                           
                        }
                     
                        //CADUtils.sele
                        
                    }

                
            }
            utils.CreateCircle(errorPC, 1, "JMD交叉");
            transaction.Commit();
        }
        }
        public static bool  PLContainsPoint(Polyline pl,Point3d pt)
        {
            Point3dCollection pc = new Point3dCollection();
            pl.GetStretchPoints(pc);
            int count = pc.Count;
            Point3d pt1 ;
            for (int a = 0; a < count; a++)
            {
                pt1 = pc[a];
                if (Math.Abs(pt1.X - pt.X) < 0.0001 && (Math.Abs(pt1.Y - pt.Y) < 0.0001))
                {
                    return true;
                }

            }
            return false;
        }

        internal void ExportRenYiTuFuKuang(string dir)
        {
            Editor editor = utils.GetEditor();
            Document document = utils.GetDocument();
             //选择指北针
            TypedValue[] types1 = new TypedValue[]
			{
				
			};
            DBObjectCollection zbzs = utils.SelectDBObjectCollection("请选择指北针", types1);
            if (zbzs.Count != 1 && zbzs[0] is BlockReference)
            {
                return;
            }
            BlockReference zbz = zbzs[0] as BlockReference;

            //1、选择要导出的图块
                //(1)JZD
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE"),
                new TypedValue((int)DxfCode.LayerName, "JZD")
			};
          
            DBObjectCollection dBObjectCollection = utils.SelectDBObjectCollection("请选择要导出的多段线", types);
            IList<Zdinfo> zdinfos = new List<Zdinfo>();
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (Entity dBObject in dBObjectCollection)
                {
                    if (dBObject is Polyline)
                    {
                    
                       
                        ResultBuffer xDataForApplication = dBObject.GetXDataForApplication("SOUTH");
                        if (!(xDataForApplication == null))
                        {
                            TypedValue[] array = xDataForApplication.AsArray();
                            if (array.Length > 3)
                            {
                                Zdinfo zdinfo = new Zdinfo();
                                zdinfo.setZdnum(array[2].Value.ToString());
                                zdinfo.Polyline = dBObject as Polyline;
                                zdinfos.Add(zdinfo);
                              
                            }
                        }                        
                    }
                }
                transaction.Commit();
                foreach (Zdinfo zdinfo in zdinfos)
                {
                    SaveZdinfoRenYiTuFuKuang(zdinfo,zbz, dir);
                   
                }
              
               
            }

            //2、循环导出每一个图块，包括所有图块内的图形
            //3、单独保存在每个文件

        }

        private void SaveZdinfoRenYiTuFuKuang(Zdinfo zdinfo, BlockReference zbz,string dir)
        {
            //添加 指北针

            string[] layerName = new string[]{
                "注记" ,"分式","面积","村","编号","用地"
            };
            PolyService polyService = new PolyService();
            DBObjectCollection saveDBC = new DBObjectCollection();
            saveDBC.Add(zdinfo.Polyline);
            //权属去包围，
            DBObjectCollection dbc = utils.GetCorssEntityCollection(zdinfo.Polyline);
            foreach (DBObject dbObject in dbc)
            {
                //只要权属线
                if (dbObject is Polyline)
                {

                    ResultBuffer xDataForApplication = dbObject.GetXDataForApplication("SOUTH");
                    if (xDataForApplication != null)
                    {
                        TypedValue[] array = xDataForApplication.AsArray();
                        if (array.Length > 3)
                        {
                            fun(saveDBC, dbObject as Polyline);
                            saveDBC.Add(dbObject);
                        }

                    }

                }
                else if (dbObject is Entity)
                {
                    Entity dbText = dbObject as Entity;

                    foreach (string str in layerName)
                    {
                        if (dbText.Layer.Contains(str))
                        {
                            saveDBC.Add(dbObject);
                            break;
                        }
                    }

                    if (dbObject is BlockReference)
                    {
                        saveDBC.Add(dbObject);
                    }
                }

            }

            //找到图形的中心点

            Point3d core = utils.GetPolyCore(zdinfo.Polyline);
            core = new Point3d(core.X - (int)(core.X % 10), core.Y - (int)(core.Y % 10), 0);
            Point3d basic = new Point3d(core.X - 150, core.Y - 100, 0);

            //计算出内矩形的左下角基准点
            //生成长300，宽200 的图形
            Point3dCollection recPoints = new Point3dCollection();
            recPoints.Add(basic);
            recPoints.Add(new Point3d(basic.X + 300, basic.Y, 0));
            recPoints.Add(new Point3d(basic.X + 300, basic.Y + 200, 0));
            recPoints.Add(new Point3d(basic.X, basic.Y + 200, 0));

            Polyline rec = polyService.Poly3dToPolyLine(utils.CreatePolye(recPoints, false));
            rec.Layer = "TK";
            addEntity(rec);
            saveDBC.Add(rec);
            ObjectId tuFutextStyle = utils.CreateSytle1("宋体", 2.4);

            {
                using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
                {
                    zbz =  transaction.GetObject(zbz.ObjectId,OpenMode.ForWrite) as BlockReference;
                    zbz = (BlockReference)zbz.Clone()  ;
                     //添加指北针
                    zbz.Position = new Point3d(recPoints[0].X+257.83, recPoints[0].Y+136.31, 0);
                    
                addEntity(zbz);
                transaction.Commit();
                saveDBC.Add(zbz);
                }
               
            }
            //填写图幅号
            {



                {
                    //左上角
                    Point3d pt = recPoints[3];

                    string xstr = (pt.X / 1000).ToString("f2");
                    DBText dBText = utils.CreateDBText(xstr, new Point3d(pt.X + 4.45 + 2.2 + 0.72, pt.Y + 8.32, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextRight;
                    dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);

                    string ystr = (pt.Y / 1000).ToString("f2");
                    dBText = utils.CreateDBText(ystr, new Point3d(pt.X - 3.66 - 1.16, pt.Y + 1.7, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextMid;
                    dBText.VerticalMode = TextVerticalMode.TextBase;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);
                }

                {
                    //右上角
                    Point3d pt = recPoints[2];

                    string xstr = (pt.X / 1000).ToString("f2");

                    DBText dBText = utils.CreateDBText(xstr, new Point3d(pt.X + 4.45, pt.Y + 8.32, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextRight;
                    dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);

                    string ystr = (pt.Y / 1000).ToString("f2");
                    dBText = utils.CreateDBText(ystr, new Point3d(pt.X + 3.66 + 1.16, pt.Y + 1.5, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextMid;
                    dBText.VerticalMode = TextVerticalMode.TextBase;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);
                }
                {
                    //右下角
                    Point3d pt = recPoints[1];

                    string xstr = (pt.X / 1000).ToString("f2");
                    DBText dBText = utils.CreateDBText(xstr, new Point3d(pt.X + 4.45, pt.Y - 8.32, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextRight;
                    dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);

                    string ystr = (pt.Y / 1000).ToString("f2");
                    dBText = utils.CreateDBText(ystr, new Point3d(pt.X + 3.66, pt.Y - 1.5, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextMid;
                    dBText.VerticalMode = TextVerticalMode.TextBase;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);
                }
                {
                    //左下角
                    Point3d pt = recPoints[0];

                    string xstr = (pt.X / 1000).ToString("f2");
                    DBText dBText = utils.CreateDBText(xstr, new Point3d(pt.X + 4.45 + 2.2, pt.Y - 8.32, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextRight;
                    dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);

                    string ystr = (pt.Y / 1000).ToString("f2");
                    dBText = utils.CreateDBText(ystr, new Point3d(pt.X - 3.66 - 1.16, pt.Y - 1.5, 0), 2.5, tuFutextStyle);
                    dBText.Layer = ("TK");
                    //dBText.ColorIndex = 1;
                    dBText.HorizontalMode = TextHorizontalMode.TextMid;
                    dBText.VerticalMode = TextVerticalMode.TextBase;
                    dBText.AlignmentPoint = dBText.Position;
                    dBText.WidthFactor = 1.0;
                    utils.AddDBText(dBText);
                    saveDBC.Add(dBText);
                }


            }



            //计算外矩形
            Point3d bigbasic = new Point3d(basic.X - 12, basic.Y - 12, 0);
            Point3dCollection bigRecPoints = new Point3dCollection();
            bigRecPoints.Add(bigbasic);
            bigRecPoints.Add(new Point3d(bigbasic.X + 324, bigbasic.Y, 0));
            bigRecPoints.Add(new Point3d(bigbasic.X + 324, bigbasic.Y + 224, 0));
            bigRecPoints.Add(new Point3d(bigbasic.X, bigbasic.Y + 224, 0));
            Polyline bigRec = polyService.Poly3dToPolyLine(utils.CreatePolye(bigRecPoints, true));
            bigRec.Layer = "TK";
            bigRec.ConstantWidth = 1;
            addEntity(bigRec);
            saveDBC.Add(bigRec);



            //添加公司文字
            ObjectId textStyle = utils.CreateSytle1("宋体", 4.0);

            Point3d start = new Point3d(bigbasic.X - 6.7, bigbasic.Y + 95.43, 0);
            string[] companyTextArray = new string[]{
                "四","川","旭","普","信","息","产","业","发","展","有","限","公","司"
             };

            for (int a = 0; a < companyTextArray.Length; a++)
            {
                start = new Point3d(start.X, start.Y - 6.3, 0);
                DBText dBText = utils.CreateDBText(companyTextArray[a], start, 4.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextCenter;
                dBText.VerticalMode = TextVerticalMode.TextBottom;
                dBText.AlignmentPoint = dBText.Position;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);
            }

            //添加文字 勘测定界图（地块105）
            {
                DBText dBText = utils.CreateDBText("勘测定界图（地块" + zdinfo.getZdnum() + ")", new Point3d(core.X + 10, core.Y + 125.73, 0), 8.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextMid;
                dBText.VerticalMode = TextVerticalMode.TextBase;
                dBText.AlignmentPoint = dBText.Position;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                //标题图号
                Point3d pt = recPoints[0];
                string xstr = (pt.X / 1000).ToString("f1");
                string ystr = (pt.Y / 1000).ToString("f1");

                ObjectId XX = utils.CreateSytle1("等线体", 2.4);
                xstr = xstr + "-" + (ystr);
                dBText = utils.CreateDBText(xstr, new Point3d(core.X + 10, core.Y + 125.73 - 9.3, 0), 5.0, XX);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextMid;
                dBText.VerticalMode = TextVerticalMode.TextBottom;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 0.75;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);
            }
            //添加 2019年12月数字化测图、2000国家大地坐标系、1985国家高程基准、2007年版图式计算机绘图

            {
                DBText dBText = utils.CreateDBText("2019年12月数字化测图", new Point3d(bigbasic.X + 12, bigbasic.Y - 5.76, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                dBText = utils.CreateDBText("2000国家大地坐标系", new Point3d(bigbasic.X + 12, bigbasic.Y - 10.76, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                dBText = utils.CreateDBText("1985国家高程基准", new Point3d(bigbasic.X + 12, bigbasic.Y - 15.76, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                dBText = utils.CreateDBText("2007年版图式计算机绘图", new Point3d(bigbasic.X + 12, bigbasic.Y - 20.76, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

            }
            {
                //添加比例
                DBText dBText = utils.CreateDBText("1：1000", new Point3d(core.X - 5.6, core.Y - 125, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextMid;
                dBText.VerticalMode = TextVerticalMode.TextBase;
                dBText.AlignmentPoint = dBText.Position;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);
            }
            {
                //添加测绘人员
                DBText dBText = utils.CreateDBText("绘图员：冯宜霞", new Point3d(core.X + 100.65, core.Y - 119.97, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                dBText = utils.CreateDBText("检查员：程婧", new Point3d(core.X + 100.65, core.Y - 124.97, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                dBText = utils.CreateDBText("审核员：吕沛橦", new Point3d(core.X + 100.65, core.Y - 129.97, 0), 3.0, textStyle);
                dBText.Layer = ("TK");
                //dBText.ColorIndex = 1;
                dBText.HorizontalMode = TextHorizontalMode.TextLeft;
                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                dBText.AlignmentPoint = dBText.Position;
                dBText.WidthFactor = 1.0;
                utils.AddDBText(dBText);
                saveDBC.Add(dBText);


            }
            {
                using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
                {

                    {
                        IList<Line> lines = new List<Line>();
                        //添加中间 直线
                        double xmin = recPoints[0].X;
                        double xMax = recPoints[1].X;
                        double ymin = recPoints[0].Y;
                        double yMax = recPoints[3].Y;

                        for (double a = ymin; a <= yMax; )
                        {
                            for (double b = xmin; b <= xMax; )
                            {
                                //生成 横 直线
                                if (b - 5 > xmin)
                                {
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(b - 5, a, 0.0), new Point3d(b + 5, a, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }
                                else
                                {
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(b, a, 0.0), new Point3d(b + 5, a, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }

                                if (a - 5 < ymin)
                                {
                                    //生成 竖 直线
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(b, a, 0.0), new Point3d(b, a + 5, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }
                                else
                                {
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(b, a - 5, 0.0), new Point3d(b, a + 5, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }


                                if (a + 100 > yMax)
                                {
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(b, yMax - 5, 0.0), new Point3d(b, yMax, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }
                                b = (int)(b / 100) * 100 + 100;
                                if (b > xMax)
                                {
                                    //添加最后一个
                                    //生成 横 直线
                                    Line line = CADUtils.CreateLine(transaction, new Point3d(xMax - 5, a, 0.0), new Point3d(xMax, a, 0.0), "TK");
                                    lines.Add(line);
                                    saveDBC.Add(line);
                                }


                            }
                            a = (int)(a / 100) * 100 + 100;

                        }
                        //添加 边缘图号文字

                        {
                            double x1 = (int)(xmin / 100) * 100 + 100;
                            double x2 = (int)(xMax / 100) * 100;
                            while (x1 <= x2)
                            {


                                //上面的文字
                                string xstr = (x1 / 1000).ToString("f2");
                                DBText dBText = utils.CreateDBText(xstr, new Point3d(x1 + 7.51, yMax + 2, 0), 2.5, tuFutextStyle);
                                dBText.Layer = ("TK");
                                dBText.HorizontalMode = TextHorizontalMode.TextRight;
                                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                                dBText.AlignmentPoint = dBText.Position;
                                dBText.WidthFactor = 1.0;
                                utils.AddDBText(dBText);
                                saveDBC.Add(dBText);
                                //下面的文字

                                xstr = (x1 / 1000).ToString("f2");
                                dBText = utils.CreateDBText(xstr, new Point3d(x1 + 7.51, ymin - 2, 0), 2.5, tuFutextStyle);
                                dBText.Layer = ("TK");
                                dBText.HorizontalMode = TextHorizontalMode.TextRight;
                                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                                dBText.AlignmentPoint = dBText.Position;
                                dBText.WidthFactor = 1.0;
                                utils.AddDBText(dBText);
                                saveDBC.Add(dBText);

                                x1 += 100;

                            }

                            double y1 = (int)(ymin / 100) * 100 + 100;
                            double y2 = (int)(yMax / 100) * 100;
                            while (y1 <= y2)
                            {


                                //上面的文字
                                string xstr = (y1 / 1000).ToString("f2");
                                DBText dBText = utils.CreateDBText(xstr, new Point3d(xmin, y1, 0), 2.3, tuFutextStyle);
                                dBText.Layer = ("TK");
                                dBText.HorizontalMode = TextHorizontalMode.TextRight;
                                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                                dBText.AlignmentPoint = dBText.Position;
                                dBText.WidthFactor = 1.0;
                                utils.AddDBText(dBText);
                                saveDBC.Add(dBText);
                                //下面的文字

                                xstr = (y1 / 1000).ToString("f2");
                                dBText = utils.CreateDBText(xstr, new Point3d(xMax + 3.5 + 7.6 + 0.2, y1, 0), 2.3, tuFutextStyle);
                                dBText.Layer = ("TK");
                                dBText.HorizontalMode = TextHorizontalMode.TextRight;
                                dBText.VerticalMode = TextVerticalMode.TextVerticalMid;
                                dBText.AlignmentPoint = dBText.Position;
                                dBText.WidthFactor = 1.0;
                                utils.AddDBText(dBText);
                                saveDBC.Add(dBText);

                                y1 += 100;

                            }



                        }
                        //添加边角八条直线
                        {
                            //左下角
                            {
                                Point3d pt = recPoints[0];
                                Line line = CADUtils.CreateLine(transaction, new Point3d(pt.X - 12, pt.Y, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);

                                line = CADUtils.CreateLine(transaction, new Point3d(pt.X, pt.Y - 12, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);
                            }
                            //右下角
                            {
                                Point3d pt = recPoints[1];
                                Line line = CADUtils.CreateLine(transaction, new Point3d(pt.X + 12, pt.Y, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);

                                line = CADUtils.CreateLine(transaction, new Point3d(pt.X, pt.Y - 12, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);
                            }
                            //右上角
                            {
                                Point3d pt = recPoints[2];
                                Line line = CADUtils.CreateLine(transaction, new Point3d(pt.X + 12, pt.Y, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);

                                line = CADUtils.CreateLine(transaction, new Point3d(pt.X, pt.Y + 12, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);
                            }
                            //左上角
                            {
                                Point3d pt = recPoints[3];
                                Line line = CADUtils.CreateLine(transaction, new Point3d(pt.X - 12, pt.Y, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);

                                line = CADUtils.CreateLine(transaction, new Point3d(pt.X, pt.Y + 12, 0.0), new Point3d(pt.X, pt.Y, 0.0), "TK");
                                saveDBC.Add(line);
                            }

                        }



                    }

                    transaction.Commit();

                }


            }
            for (int a = 0; a < saveDBC.Count; a++)
            {
                if (saveDBC[a] is Entity)
                {
                    if (((Entity)saveDBC[a]).Layer == "删除")
                    {
                        saveDBC.RemoveAt(a);
                        a--;
                    }

                }
            }

            utils.WClone(saveDBC, dir + "\\" + zdinfo.getZdnum() + ".dwg");
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject dbo in saveDBC)
                {
                    // dbo.Erase();
                }
                transaction.Commit();
            }

        }
        public void fun( DBObjectCollection saveDBC,Polyline pl)
        {
            ObjectId textStyle = utils.CreateSytle1("HZ", 4.0);
            //生成界址点号
            Point3dCollection jzds = new Point3dCollection();
            pl.GetStretchPoints(jzds);
            Point3dCollection jzdFourExPoints = new JzdService().GetJzdFourExPoints(pl);
            jzds.Add(jzds[0]);
            for (int a = 0; a < jzds.Count - 1; a++)
            {
                Point3d pt1 = jzds[a];
                Point3d pt2 = jzds[a + 1];
                double angle = CADUtils.Angle(pt1, pt2);
                Double di = utils.CalculateTwoPt(pt1, pt2);
                Point3d core;
                Point3d jzxBZ = new Point3d((pt1.X + pt2.X) / 2, (pt1.Y + pt2.Y) / 2, 0);
                if (angle >= -20 && angle < 20)
                {
                    core = new Point3d(pt1.X, pt1.Y + 6, 0);
                    if (di < 4)
                    {
                        jzxBZ = new Point3d(jzxBZ.X, jzxBZ.Y + 7, 0);
                    }
                    else
                    {
                        jzxBZ = new Point3d(jzxBZ.X, jzxBZ.Y + 6, 0);
                    }


                }
                else if (angle > -90 && angle <= 0)
                {
                    core = new Point3d(pt1.X + 6, pt1.Y, 0);
                    jzxBZ = new Point3d(jzxBZ.X + 6, jzxBZ.Y, 0);
                }
                else if (angle > -180 && angle <= -90)
                {
                    core = new Point3d(pt1.X, pt1.Y - 6, 0);
                    jzxBZ = new Point3d(jzxBZ.X - 6, jzxBZ.Y - 7, 0);
                }
                else if (angle > 20)
                {
                    core = new Point3d(pt1.X -7, pt1.Y, 0);
                    if (di < 4)
                    {
                        jzxBZ = new Point3d(jzxBZ.X - 7, jzxBZ.Y + 4.2, 0);
                    }
                    else
                    {
                        jzxBZ = new Point3d(jzxBZ.X -7, jzxBZ.Y + 4.2, 0);
                    }

                }
                //点号
                DBText dBText = utils.CreateDBText("J" + (a + 1), core, 4.0, textStyle);
                dBText.Layer = ("TK");
                dBText.ColorIndex = 1;

                dBText.HorizontalMode = TextHorizontalMode.TextMid;
                dBText.AlignmentPoint = dBText.Position;

                utils.AddDBText(dBText);
                saveDBC.Add(dBText);

                Circle c =  CADUtils.CreateCircle(jzds[a], 1);
                c.ColorIndex = 1;
                addEntity(c);
                saveDBC.Add(c);

               
            }
        }

        
        public void addEntity(Entity  db)
        {
 
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
     
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);


                blockTableRecord.AppendEntity(db);
                transaction.AddNewlyCreatedDBObject(db, true);
                transaction.Commit();
            }
           
        }
    }
}



            