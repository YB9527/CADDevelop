using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace YanBo
{
    public class Main : IExtensionApplication
    {
        private string timePath = "C:\\Windows\\System32\\123.xls";

        private static string registerDate = "2019��2��13��  12:26:58";

        private static string useDate = "2019��5��1��  22:26:58";

        private FormService formService = new FormService();

        private static UserControl1 uC;

        public string getUseDate()
        {
            return Main.useDate;
        }

        public void Initialize()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            editor.WriteMessage(">>����������ڣ�" + Main.useDate + "<<\n");
            editor.WriteMessage("����TX:����ƽ��ͼtxt�ĵ�\n");
            editor.WriteMessage("����JG:���jmd������jmd��ע��\n");
            editor.WriteMessage("����SZ:���� Ȩ������������\n");
            editor.WriteMessage("����QZB:����Ȩ����ַǩ�±�����\n");
            editor.WriteMessage("����BSB:����Ȩ����ַ��ʾ������\n");
            editor.WriteMessage("����CP:ͳ��ƽ��ͼ���������\n");
            editor.WriteMessage("����PSZ:�ڵ�ͼ������������������\n");
            editor.WriteMessage("����PA:����������塣����\n");

          
            
            Regist regist = new Regist();
            string zcm = regist.Fun();
            if (zcm != null)
            {
                editor.WriteMessage("����ϵYanBo QQ��525730167 ע�����ǣ�" + zcm);
               
            }
            else
            {
                Main.uC = new UserControl1();
                Main.uC.AddPalette();
            }
            
            /*
            string macAddressByNetworkInformation = Main.GetMacAddressByNetworkInformation();
            if (this.CheckLongTime(macAddressByNetworkInformation))
            {
                Main.uC = new UserControl1();
                Main.uC.AddPalette();
            }
            else
            {
                if (!this.Mac())
                {
                    FileStream fileStream = File.Create("D:/ע����.txt");
                    StreamWriter streamWriter = new StreamWriter(fileStream);
                    Application.ShowAlertDialog("���ɵ�ע������d�̣���ע����.txt�� ");
                    streamWriter.WriteLine(macAddressByNetworkInformation);
                    streamWriter.Close();
                    fileStream.Close();
                    HostApplicationServices.WorkingDatabase.Save();
                    for (int i = 0; i < 10000; i++)
                    {
                        Application.Quit();
                    }
                }
                if (!this.CheckTime(Main.useDate))
                {
                    Application.ShowAlertDialog("ʹ�������ѵ����������޸Ĺ���������");
                    HostApplicationServices.WorkingDatabase.Save();
                    for (int i = 0; i < 10000; i++)
                    {
                        Application.Quit();
                    }
                }
                Main.uC = new UserControl1();
                Main.uC.AddPalette();
            }*/
        }

        private bool CheckLongTime(string mac)
        {
            List<string> list = new List<string>();
            bool result;
            if (list.Contains(mac))
            {
                Main.useDate = "���ð� 12:26:58";
                result = true;
            }
            else
            {
                result = false;
            }
            return result;
        }

        public bool CheckTime(string useDate)
        {
            bool result = true;
            if (!File.Exists(this.timePath))
            {
                HSSFWorkbook hSSFWorkbook = new HSSFWorkbook();
                ISheet sheet = hSSFWorkbook.CreateSheet("Sheet1");
                FileStream fileStream = new FileStream(this.timePath, FileMode.Create);
                IRow row = sheet.CreateRow(0);
                ICell cell = row.CreateCell(0);
                cell.SetCellValue(Main.registerDate);
                row = sheet.CreateRow(1);
                cell = row.CreateCell(0);
                cell.SetCellValue(DateTime.Now.ToString("yyyy��MM��dd��  HH:mm:ss"));
                hSSFWorkbook.Write(fileStream);
                fileStream.Close();
            }
            if (Utils.DateFormat(useDate).CompareTo(DateTime.Now) > 0)
            {
                result = true;
                DateTime now = DateTime.Now;
                FileStream fileStream2 = new FileStream(this.timePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                IWorkbook workbook = new HSSFWorkbook(fileStream2);
                ISheet sheet2 = workbook[0];
                for (int i = 0; i < sheet2.LastRowNum; i++)
                {
                    IRow row2 = sheet2.GetRow(i);
                    IRow row3 = sheet2.GetRow(i + 1);
                    DateTime t = Utils.DateFormat(row2.GetCell(0).StringCellValue);
                    DateTime t2 = Utils.DateFormat(row3.GetCell(0).StringCellValue);
                    if (t > t2)
                    {
                        result = false;
                        break;
                    }
                }
                fileStream2.Close();
            }
            return result;
        }

        public void SetTime(DateTime date)
        {
            if (!File.Exists(this.timePath))
            {
                File.Create(this.timePath);
            }
            FileStream fileStream = new FileStream(this.timePath, FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(fileStream);
            ISheet sheet = hSSFWorkbook[0];
            IRow row = sheet.CreateRow(sheet.LastRowNum + 1);
            ICell cell = row.CreateCell(0);
            string cellValue = date.ToString("yyyy��MM��dd��  HH:mm:ss");
            cell.SetCellValue(cellValue);
            fileStream.Close();
            FileStream fileStream2 = new FileStream(this.timePath, FileMode.Create);
            hSSFWorkbook.Write(fileStream2);
            fileStream2.Close();
        }

        public bool Mac()
        {
            string macAddressByNetworkInformation = Main.GetMacAddressByNetworkInformation();
            Dictionary<string, string> macDic = this.GetMacDic();
            string text;
            bool result;
            if (macDic.TryGetValue(macAddressByNetworkInformation, out text))
            {
                result = true;
            }
            else
            {
                string[] array = new string[]
				{
					"EC:F4:BB:08:9A:AB",
					"1C:1B:0D:96:AD:15",
					"4C:CC:6A:D0:54:E5",
					"1C:1B:0D:99:BB:5B",
					"B0:83:FE:8F:1D:0B",
					"30:9C:23:15:4B:B4",
					"AC:22:0B:74:57:0F",
					"E0:D5:5E:4C:1D:0B"
				};
                string[] array2 = array;
                for (int i = 0; i < array2.Length; i++)
                {
                    string text2 = array2[i];
                    if (text2.Equals(macAddressByNetworkInformation))
                    {
                        result = true;
                        return result;
                    }
                }
                result = false;
            }
            return result;
        }

        private Dictionary<string, string> GetMacDic()
		{
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("30:9C:23:0F:31:3A","���´�");
			 dic.Add("1C:1B:0D:9A:52:ED","������");
             return dic;
		}

        public bool checkIp()
        {
            string cpuID = this.GetCpuID();
            string[] array = new string[]
			{
				"BFEBFBFF000906E9",
				"BFEBFBFF000306C3"
			};
            string[] array2 = array;
            bool result;
            for (int i = 0; i < array2.Length; i++)
            {
                string text = array2[i];
                if (text.Equals(cpuID))
                {
                    result = true;
                    return result;
                }
            }
            result = false;
            return result;
        }

        public void Terminate()
        {
            this.SetTime(DateTime.Now);
        }

        public static string GetMacAddressByNetworkInformation()
        {
            string str = "SYSTEM\\CurrentControlSet\\Control\\Network\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\";
            string text = string.Empty;
            try
            {
                NetworkInterface[] allNetworkInterfaces = NetworkInterface.GetAllNetworkInterfaces();
                NetworkInterface[] array = allNetworkInterfaces;
                for (int i = 0; i < array.Length; i++)
                {
                    NetworkInterface networkInterface = array[i];
                    if (networkInterface.NetworkInterfaceType == NetworkInterfaceType.Ethernet && networkInterface.GetPhysicalAddress().ToString().Length != 0)
                    {
                        string name = str + networkInterface.Id + "\\Connection";
                        RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(name, false);
                        if (registryKey != null)
                        {
                            string text2 = registryKey.GetValue("PnpInstanceID", "").ToString();
                            int num = Convert.ToInt32(registryKey.GetValue("MediaSubType", 0));
                            if (text2.Length > 3 && text2.Substring(0, 3) == "PCI")
                            {
                                text = networkInterface.GetPhysicalAddress().ToString();
                                for (int j = 1; j < 6; j++)
                                {
                                    text = text.Insert(3 * j - 1, ":");
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {
            }
            return text;
        }

        public void ReadData()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Entity entity = null;
            PromptEntityResult ent = ed.GetEntity("\nѡ��Ҫ��ȡ���ݵĶ���");
            if (ent.Status == PromptStatus.OK)
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    entity = (Entity)trans.GetObject(ent.ObjectId, OpenMode.ForRead, true);

                    ResultBuffer resBuf = entity.XData;
                    TypedValue[] tv = resBuf.AsArray();

                    if (resBuf != null)
                    {
                        //
                        IEnumerator iter = resBuf.GetEnumerator();
                        while (iter.MoveNext())
                        {
                            TypedValue tmpVal = (TypedValue)iter.Current;
                            ed.WriteMessage(tmpVal.TypeCode.ToString() + ":");
                            ed.WriteMessage(tmpVal.Value.ToString() + "\n");
                        }
                    }
                    DBDictionary NOD = (DBDictionary)trans.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead, false);


                }

            }
        }

        public string GetCpuID()
        {
            string result;
            try
            {
                ManagementClass managementClass = new ManagementClass("Win32_Processor");
                ManagementObjectCollection instances = managementClass.GetInstances();
                string text = null;
                using (ManagementObjectCollection.ManagementObjectEnumerator enumerator = instances.GetEnumerator())
                {
                    if (enumerator.MoveNext())
                    {
                        ManagementObject managementObject = (ManagementObject)enumerator.Current;
                        text = managementObject.Properties["ProcessorId"].Value.ToString();
                    }
                }
                result = text;
            }
            catch
            {
                result = "";
            }
            return result;
        }

        [CommandMethod("AD")]
        public void TextAdd()
        {
            PlanMapService planMapService = new PlanMapService();
            Utils utils = new Utils();
            IList<DBText> list = CADUtils.SeletTexts();
            if (list != null && list.Count > 1)
            {
                using (Transaction transaction = utils.GetDabase().TransactionManager.StartTransaction())
                {
                    DBText dBText = transaction.GetObject(list[0].ObjectId, OpenMode.ForWrite) as DBText;
                    string text = dBText.TextString;
                    for (int i = 1; i < list.Count; i++)
                    {
                        text = text + "��" + list[i].TextString;
                    }
                    if (dBText.Height == 1.0)
                    {
                        dBText.Height=(0.8);
                        dBText.WidthFactor=(0.8);
                    }
                    text = text.Replace("��������", "��");
                    dBText.TextString=(text);
                    transaction.Commit();
                }
            }
        }

        [CommandMethod("tx")]
        public string GetKeywords()
        {
            NFService nFService = new NFService();
            ArrayList arrayList;
            try
            {
                arrayList = nFService.readDWGNfTable();
            }
            finally
            {
                this.deleteLayerSur("ɾ��");
            }
            if (arrayList != null)
            {
                nFService.writerPMExcel(arrayList);
            }
            return "";
        }

        [CommandMethod("PSZ")]
        public string CreateZDT_SZ()
        {
            JzdService jzdService = new JzdService();
            jzdService.CreateZDT_SZ();
            return "";
        }

        [CommandMethod("cp")]
        public string CheckPlans()
        {
            try
            {
                PlanMapService planMapService = new PlanMapService();
                List<Nf> nfs = planMapService.CheckPlanArea();
                planMapService.createNfZmj(nfs);
            }
            finally
            {
                this.deleteLayerSur("ɾ��");
            }
            return "";
        }

        public void Sz()
        {
            try
            {
                JzdService jzdService = new JzdService();
                object[] array = this.formService.ShowFourAddressForm();
                if (array != null)
                {
                    jzdService.createFourAddress(array);
                }
            }
            finally
            {
                this.deleteLayerSur("ɾ��");
            }
        }

        [CommandMethod("g")]
        public void MergeGeometry()
        {
            TypedValue[] types = new TypedValue[]
			{
				new TypedValue(0, "LWPOLYLINE")
			};
            Utils utils = new Utils();
            DBObjectCollection dbc = utils.SelectDBObjectCollection("��ѡ��Ҫ�ϲ�����", types);
            CADUtils.MergeGeometrys(dbc);
        }

        [CommandMethod("qzb")]
        public string JzxTable()
        {
            try
            {
                object[] fromText = this.formService.ShowFourAddressForm();
                JzxService jzxService = new JzxService();
                List<Jzx> jzxQzb = jzxService.getJzxQzb(fromText);
                ExcelUtils excelUtils = new ExcelUtils();
                excelUtils.CreateJzxQZ_Table(jzxQzb);
            }
            finally
            {
                this.deleteLayerSur("ɾ��");
            }
            return "";
        }
        [CommandMethod("JG")]
        public DBObjectCollection Collection()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            Entity entity = null;
            DBObjectCollection EntityCollection = new DBObjectCollection();
            //���� ��������
            TypedValue value1 = new TypedValue((int)DxfCode.LayerName, "JMD");
            TypedValue[] filter = new TypedValue[1];
            filter[0] = value1;
            SelectionFilter filterLayer = new SelectionFilter(filter);

            PromptSelectionResult ents = ed.GetSelection(filterLayer);

            if (ents.Status == PromptStatus.OK)
            {
                using (Transaction transantion = db.TransactionManager.StartTransaction())
                {
                    SelectionSet SS = ents.Value;
                    foreach (ObjectId id in SS.GetObjectIds())
                    {
                        entity = (Entity)transantion.GetObject(id, OpenMode.ForRead, true);

                        if (entity != null)
                        {

                            EntityCollection.Add(entity);

                            ResultBuffer resBuf = entity.XData;
                            if (resBuf != null)
                            {
                                //
                                IEnumerator iter = resBuf.GetEnumerator();
                                while (iter.MoveNext())
                                {
                                    TypedValue tmpVal = (TypedValue)iter.Current;
                                    ed.WriteMessage(tmpVal.TypeCode.ToString() + ":");
                                    ed.WriteMessage(tmpVal.Value.ToString() + "\n");
                                }
                            }

                            setPolyeText(doc, entity);
                        }


                    }
                    transantion.Commit();
                }
            }
            return EntityCollection;

        }

        public void setPolyeText(Document doc, Entity entity)
        {
            Database database = doc.Database;
            Point3dCollection point3dCollection = new Point3dCollection();
            entity.GetStretchPoints(point3dCollection);
            Editor editor = doc.Editor;
            TypedValue typedValue = new TypedValue(8, "JMD");
            SelectionFilter selectionFilter = new SelectionFilter(new TypedValue[]
			{
				typedValue
			});
            PromptSelectionResult promptSelectionResult = editor.SelectCrossingPolygon(point3dCollection, selectionFilter);
            SelectionSet value = promptSelectionResult.Value;
            if (value != null)
            {
                editor.WriteMessage("���ݿ�ʼ");
                using (Transaction transaction = database.TransactionManager.StartTransaction())
                {
                    ObjectId[] objectIds = value.GetObjectIds();
                    for (int i = 0; i < objectIds.Length; i++)
                    {
                        ObjectId objectId = objectIds[i];
                        Entity dbText = (Entity)transaction.GetObject(objectId, OpenMode.ForWrite, true);
                        this.CheckJG(entity, dbText);
                    }
                    transaction.Commit();
                }
            }
        }

        [CommandMethod("CreateCircle")]
        public void CreateCircle(Point3d center, int colorIndex)
        {
            Database workingDatabase = HostApplicationServices.WorkingDatabase;
            using (Transaction transaction = workingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable blockTable = (BlockTable)transaction.GetObject(workingDatabase.BlockTableId, 0);
                BlockTableRecord blockTableRecord = (BlockTableRecord)transaction.GetObject(blockTable[(BlockTableRecord.ModelSpace)], OpenMode.ForWrite);
                Circle circle = new Circle(center, Vector3d.ZAxis, 10.0);
                circle.ColorIndex=(colorIndex);
                this.CreateLayer("�ṹע��");
                circle.Layer=("�ṹע��");
                blockTableRecord.AppendEntity(circle);
                transaction.AddNewlyCreatedDBObject(circle, true);
                transaction.Commit();
            }
        }

        public void CheckJG(Entity entity, Entity dbText)
        {
            if (entity != null)
            {
                Document mdiActiveDocument = Application.DocumentManager.MdiActiveDocument;
                Editor editor = mdiActiveDocument.Editor;
                ResultBuffer xData = entity.XData;
                Point3d textPoint3d = this.GetTextPoint3d(dbText);
                if (textPoint3d.X != 0.0)
                {
                    string dbText2 = this.GetDbText(dbText);
                    if (!(dbText2 == "num"))
                    {
                        if (xData != null)
                        {
                            IEnumerator enumerator = xData.GetEnumerator();
                            while (enumerator.MoveNext())
                            {
                                TypedValue typedValue = (TypedValue)enumerator.Current;
                                editor.WriteMessage(typedValue.Value.ToString() + "\n");
                                if (typedValue.Value.ToString() == "141111")
                                {
                                    editor.WriteMessage("�ṹ" + dbText2 + "\n");
                                    if (dbText2 != "��")
                                    {
                                        this.CreateCircle(textPoint3d, 1);
                                    }
                                }
                                else if (typedValue.Value.ToString() == "141121")
                                {
                                    editor.WriteMessage("�ṹ" + dbText2 + "\n");
                                    if (dbText2 != "ש")
                                    {
                                        this.CreateCircle(textPoint3d, 1);
                                    }
                                }
                                else if (typedValue.Value.ToString() == "141101")
                                {
                                    if (dbText2.Contains("ש") || dbText2.Contains("��"))
                                    {
                                        this.CreateCircle(textPoint3d, 1);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        public bool isNumber(string message)
        {
            Regex regex = new Regex("^\\d+$");
            bool result;
            if (regex.IsMatch(message))
            {
                int num = int.Parse(message);
                result = true;
            }
            else
            {
                result = false;
            }
            return result;
        }

        public string GetDbText(Entity dbText)
        {
            string result;
            if (dbText == null)
            {
                result = null;
            }
            else if (dbText is DBText)
            {
                DBText dBText = dbText as DBText;
                if (this.isNumber(dBText.TextString))
                {
                    result = "num";
                }
                else
                {
                    result = dBText.TextString;
                }
            }
            else
            {
                result = "��";
            }
            return result;
        }

        public Point3d GetTextPoint3d(Entity dbText)
        {
            Point3d result;
            if (dbText is DBText)
            {
                DBText dBText = dbText as DBText;
                result = new Point3d(dBText.Position.X, dBText.Position.Y, 0.0);
            }
            else
            {
                result = default(Point3d);
            }
            return result;
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

        [CommandMethod("BSB")]
        public string CreateJzxBsTable()
        {
            try
            {
                JzxService jzxService = new JzxService();
                List<Jzx> jzxList = jzxService.SetJzxBs();
                ExcelUtils excelUtils = new ExcelUtils();
                excelUtils.CreateJzxBS_Table(jzxList);
            }
            finally
            {
                this.deleteLayerSur("ɾ��");
            }
            return "";
        }

        public void Test()
        {
        }

        //���ɾ��ͼ���ڵ���
        public void deleteLayerSur(String layerName)
        {
            Utils utils = new Utils();
            Editor ed = utils.GetEditor();
            //��������  
            using (Transaction acTrans = utils.GetDabase().TransactionManager.StartTransaction())
            {
                //ʹ��ѡ�����������ѡ�񼯹���  
                TypedValue[] typedValue = new TypedValue[1];
                //TestLayerΪͼ������ DxfCode.LayerNameΪɸѡ����   ����������DXF����  
                typedValue.SetValue(new TypedValue((int)DxfCode.LayerName, layerName), 0);
                SelectionFilter filter = new SelectionFilter(typedValue);

                //�������� ѡ��ǰ�ռ�������δ������δ����Ķ���  
                //��ͼ����ѡ������м��ַ�ʽ��������±�  
                PromptSelectionResult result = ed.SelectAll(filter);

                // �����ʾ״̬OK����ʾ��ѡ�񵽶��� ��֮��û�ж���  
                if (result.Status != PromptStatus.OK) { return; }

                SelectionSet acSSet = result.Value;

                // ����ѡ���ڵĶ���  
                foreach (ObjectId id in acSSet.GetObjectIds())
                {
                    Entity hatchobj = acTrans.GetObject(id, OpenMode.ForWrite) as Entity;
                    hatchobj.Erase();//ɾ��  
                }
                acTrans.Commit();
            }
        }

        [CommandMethod("PA")]
        public void AddPalette()
        {
            Main.uC.AddPalette();
        }
    }
}
