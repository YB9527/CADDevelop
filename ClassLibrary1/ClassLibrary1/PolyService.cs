using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using YanBo_CG;

namespace YanBo
{
    public class PolyService
    {
        private static Utils utils = new Utils();

        public void CheckPolyRpoints()
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            DBObjectCollection dBObjectCollection = utils.SelectDBObjectCollection("请选择删除重点的多段线", types);
            using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in dBObjectCollection)
                {
                    if (dBObject is Polyline)
                    {
                        Polyline polyline = (Polyline)dBObject;
                        Point3dCollection point3dCollection = new Point3dCollection();
                        polyline.GetStretchPoints(point3dCollection);
                        Polyline polyline2 = CADUtils.PCToPolyline(PolyService.utils.RePoint(point3dCollection), polyline.GetEndWidthAt(0), polyline.Closed);
                        CADUtils.CopyAttribute(polyline, polyline2);
                        CADUtils.SaveEntity(transaction, polyline2);
                        CADUtils.DeleteEntity(transaction, dBObject.ObjectId);
                    }
                }
                transaction.Commit();
            }
        }

        public void Poly2dToPolyLine(Entity polyline2d, Entity entity)
        {
            Polyline2d polyline2d2 = (Polyline2d)entity;
            Polyline polyline = this.Poly2dTiPolyline(polyline2d, polyline2d2.DefaultStartWidth);
            polyline.Layer = entity.Layer;
            polyline.Elevation=0.0;
            this.ToModelSpace(polyline);
            polyline.XData = entity.XData;
            polyline.Color = entity.Color;
            polyline.Linetype = entity.Linetype;
            polyline.LineWeight = polyline2d.LineWeight;
           // polyline.GetPlane = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight015;
         // ConstantWidth
           polyline.ConstantWidth = 0.15;

            polyline2d.Erase();
        }

        public void Poly2dToPolyLine()
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "Polyline")
			};
            DBObjectCollection dBObjectCollection = PolyService.utils.SelectDBObjectCollection("请选择检查重点的多段线", types);
            using (Transaction transaction = PolyService.utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (Entity entity in dBObjectCollection)
                {
                    if (entity is Polyline2d)
                    {
                        Polyline2d polyline2d = (Polyline2d)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite, true);
                        this.Poly2dToPolyLine(polyline2d, entity);
                    }
                    else if (entity is Line)
                    {
                        Line polyline2d2 = (Line)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite, true);
                        this.Poly2dToPolyLine(polyline2d2, entity);
                    }
                }
                transaction.Commit();
            }
        }

        public void LineToPolyLine()
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "Line")
			};
            DBObjectCollection dBObjectCollection = PolyService.utils.SelectDBObjectCollection("请选择检查重点的多段线", types);
            using (Transaction transaction = PolyService.utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (Entity entity in dBObjectCollection)
                {
                    if (entity is Polyline2d)
                    {
                        Polyline2d polyline2d = (Polyline2d)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite, true);
                        this.Poly2dToPolyLine(polyline2d, entity);
                    }
                    else if (entity is Line)
                    {
                        Line polyline2d2 = (Line)transaction.GetObject(entity.ObjectId, OpenMode.ForWrite, true);
                        this.Poly2dToPolyLine(polyline2d2, entity);
                    }
                }
                transaction.Commit();
            }
        }

        public Polyline Poly3dToPolyLine(Polyline3d poly3d)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            poly3d.GetStretchPoints(point3dCollection);
            Point2dCollection pc = this.Point3dCollectionTo2d(point3dCollection);
            return this.Poly2dTiPolyline(pc, 0.0);
        }

        public Polyline Poly2dTiPolyline(Entity entity, double width)
        {
            Point3dCollection point3dCollection = new Point3dCollection();
            if (entity is Polyline2d)
            {
                Polyline2d polyline2d = (Polyline2d)entity;
                polyline2d.GetStretchPoints(point3dCollection);
            }
            else if (entity is Line)
            {
                Line line = (Line)entity;
                line.GetStretchPoints(point3dCollection);
            }
            return this.Poly2dTiPolyline(this.Point3dCollectionTo2d(point3dCollection), width);
        }

        public Point2dCollection Point3dCollectionTo2d(Point3dCollection pc)
        {
            Point2dCollection point2dCollection = new Point2dCollection();
            foreach (Point3d point3d in pc)
            {
                point2dCollection.Add(new Point2d(point3d.X, point3d.Y));
            }
            return point2dCollection;
        }

        public Polyline Poly2dTiPolyline(Point2dCollection pc, double width)
        {
            Polyline polyline = new Polyline();
            for (int i = 0; i < pc.Count; i++)
            {
                polyline.AddVertexAt(i, pc[i], 0.0, width, width);
            }
            polyline.Closed =true;
            return polyline;
        }

        public Polyline Poly2dTiPolyline(Point3dCollection pc, double width)
        {
            return this.Poly2dTiPolyline(this.Point3dCollectionTo2d(pc), width);
        }

        public ObjectId ToModelSpace(Entity ent)
        {
            Database dabase = PolyService.utils.GetDabase();
            ObjectId result;
            using (Transaction transaction = dabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(dabase.BlockTableId, OpenMode.ForRead);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                result = blockTableRecord.AppendEntity(ent);
                transaction.AddNewlyCreatedDBObject(ent, true);
                transaction.Commit();
            }
            return result;
        }
    }
}
