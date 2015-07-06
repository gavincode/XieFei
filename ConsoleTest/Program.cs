using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using BLL;
using ComputerInfo;

namespace ConsoleTest
{
    class Program
    {
        static String dbFile = null;

        static void Main(string[] args)
        {
            if (!CheckPermission()) return;

            Console.WriteLine("开始导入...");

#if DEBUG
             Init();

                if (dbFile == "地籍数据库")
                {
                    ExcelHandler.ExcelHandler0524.ExportTemplate();
                }
                else
                {
                    ExcelHandler.ExcelHandler.ExportTemplate();
                }

                Console.WriteLine("成功...");
                Console.WriteLine("按任意键停止并打开数据文件...");
                Console.ReadKey();
                System.Diagnostics.Process.Start("Data");
#endif

#if !DEBUG

            try
            {
                Init();

                if (dbFile == "地籍数据库")
                {
                    ExcelHandler.ExcelHandler0524.ExportTemplate();
                }
                else
                {
                    ExcelHandler.ExcelHandler.ExportTemplate();
                }

                Console.WriteLine("导出成功!");
                Console.WriteLine("按任意键停止并打开数据文件...");
                Console.ReadKey();
                System.Diagnostics.Process.Start("Data");
            }
            catch (Exception ex)
            {
                Console.WriteLine("异常信息: " + ex.Message);
                Console.WriteLine("按任意键停止...");
                Console.ReadKey();
            }

#endif
        }

        static Boolean CheckPermission()
        {
            var macAddress = new List<String>() { "48:5A:B6:92:DB:65", "48:5D:60:56:F3:B6", "74:D4:35:C9:9E:6F" };

            if (!macAddress.Contains(Computer.Instance.MacAddress))
            {
                return false;
            }

            return true;
        }

        static void Init()
        {
            var dbfiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).Where(p => p.EndsWith(".mdb"));

            if (dbfiles.Count() != 1)
            {
                throw new Exception("当前目录数据库文件不存在或不唯一！");
            }

            CommonBLL.SetDbFile(dbfiles.First());

            dbFile = Path.GetFileNameWithoutExtension(dbfiles.First());
        }
    }
}
