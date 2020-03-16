using System;
using System.Collections.Generic;
using System.ComponentModel;

using System.Data;
using System.Text;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.Runtime.InteropServices;
using System.Management;
using System.Globalization;
using Autodesk.AutoCAD.GraphicsInterface;
using YanBo.ZjdForm;
using YanBo_CG;
using System.IO;
using YanBo.office;
using NPOI.SS.UserModel;

namespace YanBo
{
    
     partial class UserControl1 : System.Windows.Forms.UserControl
    {
        private static bool psFlag = false;
        private static PaletteSet ps;
        private static Main main = new Main();
        private static Utils utils = new Utils();
        private Editor ed = utils.GetEditor();
        private PolyService polyService = new PolyService();
        private FormService formService = new FormService();
        private JzxService jzxService = new JzxService();
         private JzdService JzdService = new JzdService();

        public UserControl1()
        {
            InitializeComponent();
            //label2.Text = "�������ʱ�䣺\n" + main.getUseDate().Split(' ')[0];
            
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
        [CommandMethod("PA")]
        public void AddPalette()
        {
            
            if (psFlag == false)
            {
                psFlag = true;
                UserControl1 myControl = new UserControl1();
                //myControl.Width = 10;
                //myControl.Height = 20;
                if (ps == null)
                {
                    ps = new PaletteSet("PaletteSet");
                    ps.Style = PaletteSetStyles.ShowAutoHideButton;
                    ps.MinimumSize = new System.Drawing.Size(155, 50);
                    ps.Size = new System.Drawing.Size(155, 50);
                    ps.Add("PalettSet", myControl);
                    ps.Style = PaletteSetStyles.ShowCloseButton;
                }
                ps.Visible = true;
            }
            else
            {
                ps.Visible = false;
                psFlag = false;
            }
            
        }      
        private void button1_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.GetKeywords();
            mdiActiveDocument.Editor.WriteMessage("txt�������\n");
            documentLock.Dispose();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.Collection();
            mdiActiveDocument.Editor.WriteMessage("�ṹ��ѯ���\n");
            documentLock.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.Sz();
            mdiActiveDocument.Editor.WriteMessage("�����������\n");
            documentLock.Dispose();
        }

        private void button4_Click(object sender, EventArgs e)
        {

            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.JzxTable();
            mdiActiveDocument.Editor.WriteMessage("ǩ�±��������\n");
            documentLock.Dispose();
        }

        private void button5_Click(object sender, EventArgs e)
        {

            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.CreateJzxBsTable();
            mdiActiveDocument.Editor.WriteMessage("��ʾ���������\n");
            documentLock.Dispose();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            UserControl1.main.CheckPlans();
            mdiActiveDocument.Editor.WriteMessage("ƽ��ͼ������������\n");
            documentLock.Dispose();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            try
            {

                UserControl1.main.CreateZDT_SZ();
                mdiActiveDocument.Editor.WriteMessage("�ڵ�ͼͼ�������������\n");

            }
            finally
            {
                UserControl1.main.deleteLayerSur("ɾ��");
                documentLock.Dispose();
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            this.polyService.CheckPolyRpoints();
            mdiActiveDocument.Editor.WriteMessage("�����ɾ���ص����\n");
            documentLock.Dispose();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            this.polyService.Poly2dToPolyLine();
            mdiActiveDocument.Editor.WriteMessage("�����ת�����\n");
            documentLock.Dispose();
        }
        private void button10_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            string text = FileUtils.SaveFileName(Utils.CADDir(), "����ģ��.mdb", "Access���ݿ�(*.mdb)|*.mdb");
            if (text == null)
            {
                text = Path.GetDirectoryName(Utils.CADDir()) + "\\����ģ��.mdb";
            }
            else
            {
                FileUtils.CopyFile("C:\\ZJDTemplete\\����ģ��.mdb", text, true);
            }
            JzxService jzxService = new JzxService();
            List<Jzx> jzxList;
            try
            {
                jzxList = jzxService.SetJzxBs();
            }
            finally
            {
                UserControl1.utils.deleteLayerSur("ɾ��");
            }
            List<Jzxinfo> list = jzxService.JzxToJzxinfo(jzxList);
            List<string> sqls = OperateMdb.OpJzxinfoSql(list);
            OperateMdb.InsertSql(sqls, text);
            mdiActiveDocument.Editor.WriteMessage("mdb��ʾ���������\n");
            documentLock.Dispose();
        }
        private void button11_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            Dictionary<string, string> dictionary = ExcelUtils.ReadExcelToDic("C:\\ZJDTemplete\\��������.xls", 0);
            string text = FileUtils.SaveFileName(Utils.CADDir(), "����ģ��.mdb", "Access���ݿ�(*.mdb)|*.mdb");
            if (text == null)
            {
                text = Path.GetDirectoryName(Utils.CADDir()) + "\\����ģ��.mdb";
            }
            else
            {
                FileUtils.CopyFile("C:\\ZJDTemplete\\����ģ��.mdb", text, true);
            }
            object[] fromText = this.formService.ShowFourAddressForm();
            List<Jzx> jzxQzb = this.jzxService.getJzxQzb(fromText);
            List<Qzb> qzbList = this.jzxService.JzxToQzb(jzxQzb);
            List<string> list = OperateMdb.OpQzbSql(qzbList);
            if (list.Count != 0)
            {
                OperateMdb.InsertSql(list, text);
            }
            mdiActiveDocument.Editor.WriteMessage("mdbǩ�±��������\n");
            documentLock.Dispose();
        }
        private void installer11_AfterInstall(object sender, System.Configuration.Install.InstallEventArgs e)
        {

        }
        /**
         * //ѡ����
                Editor ed = utils.GetEditor();
                TypedValue[] types = new TypedValue[]
                {new TypedValue((int)DxfCode.LayerName,"JZD"),
                   new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                 };

                SelectionFilter sf = new SelectionFilter(types);
   
               PromptSelectionOptions options = new PromptSelectionOptions();
               options.MessageForAdding = "\n��ѡ��Ҫ����������JZDͼ��Ȩ����";
                PromptSelectionResult ents = ed.GetSelection(options,sf);
                if (ents.Status == PromptStatus.Cancel)
                {
                    return;
                }
                DBObjectCollection jzds = utils.GetEntityCollection(ents);
                if (jzds.Count == 0)
                {
                    ed.WriteMessage("\n��û��ѡ�����");
                    return;
                }
                Database db = HostApplicationServices.WorkingDatabase;

                using (Transaction trans = db.TransactionManager.StartTransaction())
                { //ѭ��ÿһ��Ȩ����
                    foreach (DBObject ob in jzds)
                    {
                        if (ob is Polyline)
                        {
                            Polyline pl = ob as Polyline;
                            Point3dCollection pts = new Point3dCollection();
                            pl.GetStretchPoints(pts);
                        
                            for (int a = 0; a < pts.Count;a++)
                            {
                                Point3d pt = pts[a];
                                Console.WriteLine("��"+a+"���㣬X:"+pt.X+"  Y:"+pt.Y);
                            }

                        }
                    }
                }
         * */
        private void button12_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            DBObjectCollection dBObjectCollection = UserControl1.utils.SelectDBObjectCollection("��Ҫ�����Ķ����", null);
            using (Transaction transaction = UserControl1.utils.GetDabase().TransactionManager.StartTransaction())
            {
                foreach (DBObject dBObject in dBObjectCollection)
                {
                    Polyline polyline = (Polyline)dBObject;
                    Point3dCollection point3dCollection = new Point3dCollection();
                    polyline.GetStretchPoints(point3dCollection);
                    Point3d point3d = point3dCollection[0];
                }
                transaction.Commit();
            }
            documentLock.Dispose();

        }

        private void button13_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("���ֶ��ƿ�������ϵ�ˣ�YanBo����ϵ��ʽ��QQ:525730167");
        }

        private void button14_Click(object sender, EventArgs e)
        {


            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            try
            {
                PlanMapForm planMapForm = new PlanMapForm();
                planMapForm.ShowDialog();
            }
            finally
            {
                UserControl1.main.deleteLayerSur("ɾ��");
            }
            documentLock.Dispose();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
            DocumentLock documentLock = mdiActiveDocument.LockDocument();
            this.polyService.LineToPolyLine();
            mdiActiveDocument.Editor.WriteMessage("�����ת�����\n");
            documentLock.Dispose();
        }

         private void button16_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             UserControl1.main.MergeGeometry();
             documentLock.Dispose();
         }
         /*
         private void button17_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
             DBObjectCollection dbc = UserControl1.utils.SelectDBObjectCollection("��Ҫ�����Ķ����", types);
             CADUtils.Adjustment(dbc);
             documentLock.Dispose();
         }*/
        

         private void button19_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 PlanMapService planMapService = new PlanMapService();
                 planMapService.FlushPlanMap();
             }
             finally
             {
                 documentLock.Dispose();
             }
         }

         private void button20_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 PlanMapService planMapService = new PlanMapService();
                 planMapService.DeleteZdnum2();
             }
             finally
             {
                 documentLock.Dispose();
             }
         }

         private void button21_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 this.JzdService.ExcportJZDXLX();
             }
             finally
             {
                 documentLock.Dispose();
             }
         }

         private void button22_Click_1(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 string path = Utils.CADDir();
                 string text = FileUtils.SaveFileName(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + "��ַ��ɹ���", "Excel files(*.xls)|*.xls");
                 if (text != null)
                 {
                     this.JzdService.ExportJzdTable(text);
                 }
             }
             finally
             {
                 documentLock.Dispose();
             }
         }

         private void button23_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             string text = FileUtils.SaveFileName(Path.GetDirectoryName(Utils.CADDir()), "����ģ��.mdb", "Access���ݿ�(*.mdb)|*.mdb");
             if (text == null)
             {
                 text = Path.GetDirectoryName(Utils.CADDir()) + "\\����ģ��.mdb";
             }
             else
             {
                 FileUtils.CopyFile("C:\\ZJDTemplete\\����ģ��.mdb", text, true);
             }
             JzxService jzxService = new JzxService();
             List<Jzx> jzxList;
             try
             {
                 jzxList = jzxService.SetJzxBs();
             }
             finally
             {
                 UserControl1.utils.deleteLayerSur("ɾ��");
             }
             List<Jzxinfo> list = jzxService.JzxToJzxinfo(jzxList);
             List<string> list2 = OperateMdb.OpJzxinfoSql(list);
             OperateMdb.InsertSql(list2, text);
             mdiActiveDocument.Editor.WriteMessage("mdb��ʾ���������\n");
             object[] fromText = this.formService.ShowFourAddressForm();
             List<Jzx> jzxQzb = jzxService.getJzxQzb(fromText);
             List<Qzb> qzbList = jzxService.JzxToQzb(jzxQzb);
             List<string> sqls = OperateMdb.OpQzbSql(qzbList);
             if (list2.Count != 0)
             {
                 OperateMdb.InsertSql(sqls, text);
             }
             mdiActiveDocument.Editor.WriteMessage("mdbǩ�±��������\n");
             documentLock.Dispose();
         }

         private void button24_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                
                 string text = FileUtils.SelectSingleExcelFile("ѡ�����ڵ����Ա�");
                 string path = Utils.CADDir();
                 string text2 = FileUtils.SeleFileDir(Path.GetDirectoryName(path));
                 if (text2 != null || text != null)
                 {
                     JzdService.ExportFWDCB(text2, text);
                 }
             }
             finally
             {
                 main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }


         private void button26_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 TypedValue[] types = new TypedValue[]
				{
					new TypedValue(0, "LWPOLYLINE")
				};
                 DBObjectCollection dBObjectCollection = UserControl1.utils.SelectDBObjectCollection("ѡ����Ҫͳ�ƽ�ַ�������Ȩ����", types);
                 if (dBObjectCollection != null && dBObjectCollection.Count != 0)
                 {
                     string path = Utils.CADDir();
                     NPOI.SS.UserModel.IWorkbook workbook = new JzdService().ExportJzdsExcel(dBObjectCollection);
                     if (workbook == null)
                     {
                         this.ed.WriteMessage("��û��ѡ�����ߵ�Ȩ���ߣ�����");
                     }
                     else
                     {
                         string text = FileUtils.SaveFileName(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + "�ؼ�ͼ��ַ�����ͳ��", FileUtils.ExcelFilter);
                         if (text != null)
                         {
                             ExcelWrite.Save(workbook, text);
                         }
                     }
                 }
             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }

         private void button27_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 string path = Utils.CADDir();
                 IWorkbook workbook = new JzdService().ZDTExportJzdsExcel();
                 if (workbook != null)
                 {
                     string text = FileUtils.SaveFileName(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + "�ڵ�ͼ��ַ�����ͳ��", FileUtils.ExcelFilter);
                     if (text != null)
                     {
                         ExcelWrite.Save(workbook, text);
                     }
                 }
             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }

         private void button25_Click_1(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 string path = Utils.CADDir();
                 string text = FileUtils.SeleFileDir("Ȩ���������λ��", Path.GetDirectoryName(path));
                 if (text != null)
                 {
                     new JzdService().ExportQJDCB(text);
                 }
             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }
         /// <summary>
         /// ��齨������
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void button15_Click_1(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                //���ݼ�������
                //new PlanMapService().CheckJZWContains();

                 new PlanMapService().CheckJZWContains2();

                //Utils utils = new Utils();
                //utils.CreateCircle(new Point3d(0, 0, 0), 1, "0");
             }
             finally
             {

                 documentLock.Dispose();
             }
         }

         /// <summary>
         /// ����ڵ�ͼ�����������ͼ���š�
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void button17_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {

                 new PlanMapService().CheckZDT();
              
             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }
         /// <summary>
         ///  �ڵ�ͼ����1
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void button18_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {
                 PlanMapService planMapService = new PlanMapService();
                 planMapService.RegulatePlanMap();
             }
             finally
             {
                 documentLock.Dispose();
             }
         }
        /// <summary>
        /// �ڵ�ͼ����2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
         private void button27_Click_1(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {

                 PlanMapService planMapService = new PlanMapService();
                 planMapService.RegulatePlanMap2();
             

             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
          
         }

         /// <summary>
         /// �ڵ�ͼ������ʽ����
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void button28_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             try
             {

                 PlanMapService planMapService = new PlanMapService();
                 planMapService.ZDTTextManger();


             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }

         }

         /// <summary>
         /// ��������ͼ��
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
         private void button29_Click(object sender, EventArgs e)
         {
             Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
             DocumentLock documentLock = mdiActiveDocument.LockDocument();
             string path = Utils.CADDir();
             //string text = FileUtils.SeleFileDir("Ȩ���������λ��", Path.GetDirectoryName(path));
             try
             {

                 PlanMapService planMapService = new PlanMapService();
                 planMapService.ExportRenYiTuFuKuang(Path.GetDirectoryName(path));


             }
             finally
             {
                 UserControl1.main.deleteLayerSur("ɾ��");
                 documentLock.Dispose();
             }
         }
        
    }
}
