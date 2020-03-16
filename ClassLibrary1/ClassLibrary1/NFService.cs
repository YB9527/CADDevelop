using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections;
using System.IO;
using YanBo_CG;

namespace YanBo
{
    internal class NFService
    {
        private Utils utils = new Utils();

        public void writerPMExcel(ArrayList list)
        {
            FileStream fileStream = File.Create("D:/test.txt");
            StreamWriter streamWriter = new StreamWriter(fileStream);
            foreach (Nf nf in list)
            {
                string zdnum = nf.getZdnum();
                string cbfmc = nf.getCbfmc();
                string cbfdz = nf.getCbfdz();
                string zmj = nf.getZmj();
                string czs = nf.getCzs();
                string[] hhMj = nf.getHhMj();
                string[] zmMj = nf.getZmMj();
                string[] qtMj = nf.getQtMj();
                zmj = nf.getZmj();
                czs = nf.getCzs();
                string value = nf.toString();
                streamWriter.WriteLine(value);
            }
            streamWriter.Close();
            fileStream.Close();
        }

        public ArrayList readDWGNfTable()
        {
            ArrayList arrayList = new ArrayList();
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            Utils utils = new Utils();
            string textString = utils.getTextEntity("请选择房产平面图单个图框标题文字").TextString;
            ArrayList result;
            if (textString == null)
            {
                editor.WriteMessage("你选择的不是单个的文字");
                result = null;
            }
            else
            {
                Point3dCollection point3dCollection = utils.findEntityPoints(textString);
                TypedValue[] array = new TypedValue[]
				{
					new TypedValue(0, "LINE")
				};
                Point3dCollection point3dCollection2 = new Point3dCollection();
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
                    PromptSelectionResult ents = editor.SelectCrossingPolygon(point3dCollection4);
                    DBObjectCollection entityCollection = this.GetEntityCollection(ents);
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
                foreach (Point3d point3d in point3dCollection2)
                {
                    Nf nf = new Nf();
                    Point3d point3d5 = new Point3d(point3d.X + 4.54, point3d.Y, 0.0);
                    Point3d pt = new Point3d(point3d.X + 12.87, point3d.Y - 2.86, 0.0);
                    string windowText = this.getWindowText(point3d5, pt);
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
                    point3d5 = new Point3d(point3d.X + 4.54, point3d.Y - 2.86, 0.0);
                    pt = new Point3d(point3d.X + 12.87, point3d.Y - 5.72, 0.0);
                    windowText = this.getWindowText(point3d5, pt);
                    nf.setCbfmc(windowText);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage("编码是：" + nf.getZdnum() + ", 名字没有填写 ");
                    }
                    point3d5 = new Point3d(point3d.X + 10.72, point3d.Y - 66.55, 0.0);
                    pt = new Point3d(point3d.X + 42.9, point3d.Y - 68.95, 0.0);
                    windowText = this.getWindowText(point3d5, pt);
                    nf.setCbfdz(windowText);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage("编码是：" + nf.getZdnum() + ", 房屋坐落没有填写 ");
                    }
                    point3d5 = new Point3d(point3d.X + 10.72, point3d.Y - 68.95, 0.0);
                    pt = new Point3d(point3d.X + 21.45, point3d.Y - 71.35, 0.0);
                    windowText = this.getWindowText(point3d5, pt);
                    nf.setCzs(windowText);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage("编码是：" + nf.getZdnum() + ", 层数没有填写 ");
                    }
                    point3d5 = new Point3d(point3d.X + 32.17, point3d.Y - 68.95, 0.0);
                    pt = new Point3d(point3d.X + 42.9, point3d.Y - 71.35, 0.0);
                    windowText = this.getWindowText(point3d5, pt);
                    if (windowText == null || "".Equals(windowText))
                    {
                        editor.WriteMessage("编码是：" + nf.getZdnum() + ", 总面积没有填写 ");
                    }
                    nf.setZmj(windowText);
                    string[] hhmj = new string[3];
                    string[] zmmj = new string[3];
                    string[] qtmj = new string[3];
                    point3d5 = new Point3d(point3d.X, point3d.Y - 71.35, 0.0);
                    pt = new Point3d(point3d.X + 10.72, point3d.Y - 73.75, 0.0);
                    string windowText2 = this.getWindowText(point3d5, pt);
                    this.setNfCS(nf, point3d5, windowText2, 0, hhmj, zmmj, qtmj);
                    point3d5 = new Point3d(point3d.X + 10.72, point3d.Y - 71.35, 0.0);
                    pt = new Point3d(point3d.X + 21.45, point3d.Y - 73.75, 0.0);
                    windowText2 = this.getWindowText(point3d5, pt);
                    this.setNfCS(nf, point3d5, windowText2, 1, hhmj, zmmj, qtmj);
                    point3d5 = new Point3d(point3d.X + 21.45, point3d.Y - 71.35, 0.0);
                    pt = new Point3d(point3d.X + 32.17, point3d.Y - 73.75, 0.0);
                    windowText2 = this.getWindowText(point3d5, pt);
                    this.setNfCS(nf, point3d5, windowText2, 2, hhmj, zmmj, qtmj);
                    string[] sub = new string[3];
                    point3d5 = new Point3d(point3d.X + 32.17, point3d.Y - 71.35, 0.0);
                    pt = new Point3d(point3d.X + 42.9, point3d.Y - 73.75, 0.0);
                    windowText2 = this.getWindowText(point3d5, pt);
                    this.setSubTotal(nf, point3d5, sub);
                    point3d5 = new Point3d(point3d.X + 50.05, point3d.Y - 76.15, 0.0);
                    pt = new Point3d(point3d.X + 64.35, point3d.Y - 78.55, 0.0);
                    string windowText3 = this.getWindowText(point3d5, pt);
                    nf.setYear(windowText3);
                    arrayList.Add(nf);
                }
                result = arrayList;
            }
            return result;
        }

        public void setSubTotal(Nf nf, Point3d position, string[] sub)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            for (int i = 0; i < 5; i++)
            {
                string text = "";
                for (int j = 0; j < 2; j++)
                {
                    double num = position.X + (double)j * 5.36;
                    double num2 = position.Y - (double)(i + 2) * 2.4;
                    double num3 = position.X + (double)(j + 1) * 5.36;
                    double num4 = position.Y - (double)(i + 3) * 2.4;
                    Point3d pt = new Point3d(num, num2, 0.0);
                    Point3d pt2 = new Point3d(num3, num4, 0.0);
                    string windowText = this.getWindowText(pt, pt2);
                    if (windowText.Trim().Equals(""))
                    {
                        break;
                    }
                    if (j == 0)
                    {
                        text = text + "小计" + windowText + ":";
                    }
                    else
                    {
                        text = text + "" + windowText;
                    }
                    if (text.Contains("混合") && text.Length > 3 && !windowText.Equals(""))
                    {
                        if (sub[0] != null)
                        {
                        }
                        sub[0] = text;
                    }
                    else if (text.Contains("砖木") && text.Length > 3 && !windowText.Equals(""))
                    {
                        if (sub[1] != null)
                        {
                        }
                        sub[1] = text;
                    }
                    else if (text.Contains("其他") && text.Length > 3 && !windowText.Equals(""))
                    {
                        if (sub[2] != null)
                        {
                        }
                        sub[2] = text;
                    }
                    else if (!text.Contains("小计混合:") && !text.Contains("小计砖木:") && !text.Contains("小计其他:") && !windowText.Equals(""))
                    {
                        editor.WriteMessage(string.Concat(new string[]
						{
							"姓名是：",
							nf.getCbfmc(),
							"结构:",
							text,
							",有问题"
						}));
                    }
                }
            }
            nf.setSubTotal(sub);
        }

        public void setNfCS(Nf nf, Point3d position, string cs, int index, string[] hhmj, string[] zmmj, string[] qtmj)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            for (int i = 0; i < 5; i++)
            {
                string text = "";
                for (int j = 0; j < 2; j++)
                {
                    double num = position.X + (double)j * 5.36;
                    double num2 = position.Y - (double)(i + 2) * 2.4;
                    double num3 = position.X + (double)(j + 1) * 5.36;
                    double num4 = position.Y - (double)(i + 3) * 2.4;
                    Point3d pt = new Point3d(num, num2, 0.0);
                    Point3d pt2 = new Point3d(num3, num4, 0.0);
                    string windowText = this.getWindowText(pt, pt2);
                    if (windowText.Trim().Equals(""))
                    {
                        break;
                    }
                    if (j == 0)
                    {
                        text = cs + ":" + windowText;
                    }
                    else
                    {
                        text = text + ":" + windowText;
                    }
                    if (text.Contains("混合") && text.Length > 5 && !windowText.Equals(""))
                    {
                        if (hhmj[index] != null)
                        {
                            editor.WriteMessage(string.Concat(new string[]
							{
								"编码是：",
								nf.getZdnum(),
								",",
								cs,
								" 层结构有问题"
							}));
                        }
                        hhmj[index] = text;
                    }
                    else if (text.Contains("砖木") && text.Length > 5 && !windowText.Equals(""))
                    {
                        if (zmmj[index] != null)
                        {
                            editor.WriteMessage(string.Concat(new string[]
							{
								"编码是：",
								nf.getZdnum(),
								",",
								cs,
								" 层结构有问题"
							}));
                        }
                        zmmj[index] = text;
                    }
                    else if (text.Contains("其他") && text.Length > 5 && !windowText.Equals(""))
                    {
                        if (qtmj[index] != null)
                        {
                            editor.WriteMessage(string.Concat(new string[]
							{
								"编码是：",
								nf.getZdnum(),
								",",
								cs,
								" 层结构有问题"
							}));
                        }
                        qtmj[index] = text;
                    }
                    else if (!text.Contains(":混合") && !text.Contains(":砖木") && !text.Contains(":其他") && !windowText.Equals(""))
                    {
                        editor.WriteMessage(string.Concat(new string[]
						{
							"姓名是：",
							nf.getCbfmc(),
							"结构:",
							text,
							",有问题"
						}));
                    }
                }
            }
            nf.setHhMj(hhmj);
            nf.setZmMj(zmmj);
            nf.setQtMj(qtmj);
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
            Polyline3d polyline3d = this.CreatePolye(point3dCollection);
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

        public Point3dCollection CirclesPoints()
        {

            NFService nfServies = new NFService();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //得到用户输入值
            //得到
            PromptIntegerOptions op = new PromptIntegerOptions("输入基点圆的半径");
            int value = 10;
            PromptIntegerResult radius;
            radius = ed.GetInteger(op);
            if (radius.Status == PromptStatus.OK)
            {
                value = radius.Value;
            }
            DBObjectCollection cilres = nfServies.SelectCircel(value);
            Point3dCollection points = new Point3dCollection();

            //循环圆 得到圆心
            DBObjectCollection txts = nfServies.SelectCircel(value);
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                foreach (DBObject ob in cilres)
                {
                    Circle circle = (Circle)trans.GetObject(ob.ObjectId, OpenMode.ForWrite, true);
                    points.Add(circle.Center);
                }
                trans.Commit();
            }
            return points;
        }
    
        public DBObjectCollection SelectCircel(int radius)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            TypedValue[] array = new TypedValue[]
			{
				new TypedValue(0, "CIRCLE"),
				new TypedValue(40, radius)
			};
            SelectionFilter selectionFilter = new SelectionFilter(array);
            PromptSelectionResult selection = editor.GetSelection(selectionFilter);
            return this.GetEntityCollection(selection);
        }

        public DBObjectCollection GetEntityCollection(PromptSelectionResult ents)
        {

            Database db = HostApplicationServices.WorkingDatabase;
            Entity entity = null;
            DBObjectCollection EntityCollection = new DBObjectCollection();
            if (ents.Status == PromptStatus.OK)
            {
                using (Transaction transaction = db.TransactionManager.StartTransaction())
                {
                    SelectionSet ss = ents.Value;
                    foreach (ObjectId id in ss.GetObjectIds())
                    {
                        entity = transaction.GetObject(id, OpenMode.ForWrite, true) as Entity;
                        if (entity != null)
                        {
                            EntityCollection.Add(entity);
                        }
                    }
                    transaction.Commit();
                }
            }
            return EntityCollection;

        }
        public double Di(Point3d pt1, Point3d pt2)
        {
            double num = pt1.X - pt2.X;
            double num2 = pt1.Y - pt2.Y;
            return Math.Abs(Math.Sqrt(num * num + num2 * num2));
        }

        public Polyline3d CreatePolye(Point3dCollection points)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Polyline3d polyline3d = new Polyline3d(0, points, true);
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, 0);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                polyline3d.ColorIndex=(1);
                this.CreateLayer("删除");
                polyline3d.Layer=("删除");
                blockTableRecord.AppendEntity(polyline3d);
                transaction.AddNewlyCreatedDBObject(polyline3d, true);
                transaction.Commit();
            }
            return polyline3d;
        }

        public void CreateLayer(string layer)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            Transaction transaction = workingDatabase.TransactionManager.StartTransaction();
            try
            {
                LayerTable layerTable = (LayerTable)transaction.GetObject(workingDatabase.LayerTableId, OpenMode.ForWrite);
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
    }
}
