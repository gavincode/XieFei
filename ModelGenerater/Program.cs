using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace ModelGenerater
{
    class Program
    {
        static void Main(string[] args)
        {
            var dbFile = GetDbFile();

            foreach (var tableName in GetTableNames(dbFile))
            {
                var table = ReadTable(dbFile, tableName);
                table.TableName = tableName;

                var sql = GenerateModel(table);

                File.WriteAllText(tableName + ".cs", sql);
            }

            Console.WriteLine("生成成功! 任意键退出...");
            Console.ReadKey();
            System.Diagnostics.Process.Start(Environment.CurrentDirectory);
        }

        static String GetDbFile()
        {
            var dbfiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).Where(p => p.EndsWith(".mdb"));

            if (dbfiles.Count() != 1)
            {
                throw new Exception("当前目录数据库文件不存在或不唯一！");
            }

            return dbfiles.First();
        }

        static OleDbConnection getConn(String dbFile)
        {
            string connstr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + dbFile;
            OleDbConnection tempconn = new OleDbConnection(connstr);
            return (tempconn);
        }

        static IEnumerable<String> GetTableNames(String dbFile)
        {
            using (var conn = getConn(dbFile))
            {
                conn.Open();
                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                return dt.AsEnumerable().Select(row => row.Field<String>("TABLE_NAME")).ToList();
            }
        }

        static DataTable ReadTable(String dbFile, String tableName)
        {
            using (var conn = getConn(dbFile))
            {
                var sql = String.Format("select * from {0};", tableName);

                OleDbCommand cmd = new OleDbCommand(sql, conn);
                OleDbDataAdapter daper = new OleDbDataAdapter(cmd);
                DataSet dataSet = new DataSet();
                daper.Fill(dataSet);

                return dataSet.Tables[0];
            }
        }

        static String GenerateModel(DataTable table)
        {
            var mainSQL = @"using System;
                            using System.Collections.Generic;
                            using System.Data;
                            using System.Linq;

                            namespace Model
                            {
                                public class @ModelName
                                {@Properties
                                }
                            }
                                                        ";

            String properties = null;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                properties += String.Format(@"{0}public String {1} {2} get; set; {3}", Environment.NewLine, table.Columns[i].ColumnName, "{", "}");
            }

            return mainSQL.Replace("@ModelName", table.TableName).Replace("@Properties", properties);
        }
    }
}
