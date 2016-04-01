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

                Console.WriteLine("成功...");
                System.Diagnostics.Process.Start("Data");
            }
            catch (Exception ex)
            {
                Console.WriteLine("异常信息: " + ex.Message);
                Console.WriteLine("异常信息: " + ex.StackTrace);
                Console.WriteLine("按任意键停止...");
                Console.ReadKey();
            }
        }

        static Boolean CheckPermission()
        {
            var macAddress = new List<String>() { "34:E6:AD:0A:8F:AC", "54:04:A6:93:FF:0F", "48:5A:B6:92:DB:65", "48:5D:60:56:F3:B6", "AC:D1:B8:48:AF:53", "BFEBFBFF000306A9" };

            if (!macAddress.Contains(Computer.Instance.MacAddress))
            {
                return false;
            }

            return true;
        }

        static void Init()
        {
            var dbfiles = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).Where(p => p.ToLower().EndsWith(".mdb"));

            if (dbfiles.Count() != 1)
            {
                throw new Exception("当前目录数据库文件不存在或不唯一！");
            }

            CommonBLL.SetDbFile(dbfiles.First());

            dbFile = Path.GetFileNameWithoutExtension(dbfiles.First());
        }
    }
}
