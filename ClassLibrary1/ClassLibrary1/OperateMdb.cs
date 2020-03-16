using ADOX;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace YanBo
{
    internal class OperateMdb
    {
        public static Dictionary<string, List<Jzxinfo>> MdbToJzxinfo(string mdbPath)
        {
            DataTable dataTable = OperateMdb.ReadAllData("jzxinfo", mdbPath, true);
            DataRowCollection rows = dataTable.Rows;
            List<Jzxinfo> list = null;
            Dictionary<string, List<Jzxinfo>> dictionary = new Dictionary<string, List<Jzxinfo>>();
            foreach (DataRow dataRow in rows)
            {
                object[] itemArray = dataRow.ItemArray;
                Jzxinfo jzxinfo = new Jzxinfo();
                int i = 0;
                while (i < itemArray.Length)
                {
                    switch (i)
                    {
                        case 0:
                            jzxinfo.setBzdh((string)itemArray[i]);
                            break;
                        case 2:
                            jzxinfo.setQdh((string)itemArray[i]);
                            break;
                        case 3:
                            jzxinfo.setZdh((string)itemArray[i]);
                            break;
                        case 4:
                            {
                                double tsbc = 0.0;
                                try
                                {
                                    tsbc = double.Parse((string)itemArray[i]);
                                }
                                catch
                                {
                                }
                                jzxinfo.setTsbc(tsbc);
                                break;
                            }
                    }
                IL_E4:
                    i++;
                    continue;
                    goto IL_E4;
                }
                if (jzxinfo.getTsbc() > 0.0)
                {dictionary.TryGetValue(jzxinfo.getBzdh(), out list);
                if (list == null)
                {
                    list = new List<Jzxinfo>();
                    list.Add(jzxinfo);
                    dictionary.Add(jzxinfo.getBzdh(), list);
                }
                else
                {
                    list.Add(jzxinfo);
                }
                   
                }

            }
            return dictionary;
        }

        public static DataTable ReadAllData(string tableName, string mdbPath, bool success)
        {
            DataTable dataTable = new DataTable();
            DataTable result;
            try
            {
                string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath;
                OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
                oleDbConnection.Open();
                OleDbCommand oleDbCommand = oleDbConnection.CreateCommand();
                oleDbCommand.CommandText = "select * from " + tableName;
                OleDbDataReader oleDbDataReader = oleDbCommand.ExecuteReader();
                int fieldCount = oleDbDataReader.FieldCount;
                for (int i = 0; i < fieldCount; i++)
                {
                    DataColumn column = new DataColumn(oleDbDataReader.GetName(i));
                    dataTable.Columns.Add(column);
                }
                while (oleDbDataReader.Read())
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int i = 0; i < fieldCount; i++)
                    {
                        dataRow[oleDbDataReader.GetName(i)] = oleDbDataReader[oleDbDataReader.GetName(i)].ToString();
                    }
                    dataTable.Rows.Add(dataRow);
                }
                oleDbDataReader.Close();
                oleDbConnection.Close();
                result = dataTable;
            }
            catch
            {
                result = dataTable;
            }
            return result;
        }

        public static bool CreateMDBDataBase(string mdbPath)
        {
            bool result;
            try
            {
                CatalogClass catalogClass = new CatalogClass();
                catalogClass.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbPath + ";");
                result = true;
            }
            catch
            {
                result = false;
            }
            return result;
        }

        public static void InsertSql(List<string> sqls, string mdbPath)
        {
            OleDbConnection oleDbConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + mdbPath);
            oleDbConnection.Open();
            foreach (string current in sqls)
            {
                OleDbCommand oleDbCommand = new OleDbCommand(current, oleDbConnection);
                int num = oleDbCommand.ExecuteNonQuery();
            }
            oleDbConnection.Close();
        }

        public static List<string> OpJzxinfoSql(List<Jzxinfo> list)
        {
            List<string> list2 = new List<string>();
            foreach (Jzxinfo current in list)
            {
                string bzdh = current.getBzdh();
                string lzdh = current.getLzdh();
                string qdh = current.getQdh();
                string zdh = current.getZdh();
                double tsbc = current.getTsbc();
                double kzbc = current.getKzbc();
                string jxxz = current.getJxxz();
                string jzxlb = current.getJzxlb();
                string jzxwz = current.getJzxwz();
                string bzdzjr = current.getBzdzjr();
                string bzdzjr2 = current.getBzdzjr();
                string lzdzjr = current.getLzdzjr();
                string lzdzjrq = current.getLzdzjrq();
                string item = string.Concat(new object[]
				{
					"insert into jzxinfo(BZDH,LZDH,QDH,ZDH,TSBC,KZBC,JXXZ,JZXLB,JZXWZ,BZDZJR,BZDZJRQ,LZDZJR,LZDZJRQ,QDZB,ZDZB) values ('",
					bzdh,
					"','",
					lzdh,
					"','",
					qdh,
					"','",
					zdh,
					"','",
					tsbc,
					"','",
					kzbc,
					"','",
					jxxz,
					"','",
					jzxlb,
					"','",
					jzxwz,
					"','",
					bzdzjr,
					"','",
					bzdzjr2,
					"','",
					lzdzjr,
					"','",
					lzdzjrq,
					"','",
					current.getQdzb(),
					"','",
					current.getZdzb(),
					"')"
				});
                list2.Add(item);
            }
            return list2;
        }

        public static List<string> OpQzbSql(List<Qzb> qzbList)
        {
            List<string> list = new List<string>();
            foreach (Qzb current in qzbList)
            {
                string item = string.Concat(new string[]
				{
					"insert into qzb(BZDH, LZDH, QDH, ZDH, LZDZJR) values('",
					current.getBzdh(),
					"','",
					current.getLzdh(),
					"','",
					current.getQdh(),
					"','",
					current.getZdh(),
					"','",
					current.getLzdzjr(),
					"')"
				});
                list.Add(item);
            }
            return list;
        }
    }
}
