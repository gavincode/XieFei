using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using BLL;
using Excel.Export.AutoCode;
using Model;
using NPOI.SS.UserModel;

namespace ExcelHandler
{
    public class ExcelHandler
    {
        const String templateFile = @"Template\登记表.xls";
        const String file = @"Data\{0}.xls";

        public static void Init()
        {
            if (!Directory.Exists("Data"))
                Directory.CreateDirectory("Data");
            if (!File.Exists(templateFile))
                Console.WriteLine("模板不存在: " + templateFile);
        }

        public static void ExportTemplate()
        {
            Init();

            var 承包方调查表 = CommonBLL.承包方调查表();
            var 二轮承包表 = CommonBLL.二轮承包表();
            var 地块属性表 = CommonBLL.地块属性表();

            foreach (var dkGroup in 地块属性表)
            {
                登记表.Workbook = WorkbookFactory.Create(templateFile);

                var first地块 = dkGroup.Value.First();

                if (!承包方调查表.ContainsKey(first地块.承包方代表编码))
                {
                    continue;
                    throw new Exception(String.Format("承包方调查表中不存在承包方编码为{0}的数据。", first地块.承包方代表编码));
                }

                var 承包方 = 承包方调查表[first地块.承包方代表编码];

                var 二轮承包信息 = new List<二轮承包表>();
                if (二轮承包表.ContainsKey(first地块.承包方代表编码))
                {
                    二轮承包信息 = 二轮承包表[first地块.承包方代表编码];
                }

                //封面
                登记表.封面.承包方代表姓名.Fill(first地块.承包方代表姓名);
                登记表.封面.个体编码.Fill(first地块.个体编码);
                登记表.封面.行政区名称.Fill(first地块.行政区名称);

                //登记簿1
                登记表.登记簿1.承包方代表姓名.Fill(承包方.承包方名称);
                登记表.登记簿1.承包方地址.Fill(承包方.承包方地址);
                登记表.登记簿1.个体编码.Fill(first地块.个体编码);
                登记表.登记簿1.个体编码1.Fill(first地块.个体编码);
                登记表.登记簿1.行政区名称.Fill(first地块.行政区名称);
                登记表.登记簿1.邮编.Fill(承包方.邮政编码);
                登记表.登记簿1.证件号码.Fill(承包方.证件号码);
                登记表.登记簿1.非在册户口.Fill(承包方.非在册户口);

                登记表.登记簿1.Row11.Fill(承包方.所属家属列表, p => new List<Object>
                {
                    p.成员姓名,
                    p.家庭关系,
                    p.证件号码,
                    p.有承包地,
                    p.无承包地,
                    ToNumber(p.备注 == "在册", 1), //? 1 : 0,
                    ToNumber(p.备注 == "死亡", 1),  // : 0,
                    ToNumber(p.备注 == "迁出", 1), //: 0,
                    //p.备注
                });

                //登记簿2
                var groupOther = dkGroup.Value.Where(p => (p.地块类别 ?? "").ToLower() != "z");
                var groupZ = dkGroup.Value.Where(p => (p.地块类别 ?? "").ToLower() == "z");

                Func<地块属性表, List<Object>> func = p => new List<Object>
                {
                    p.地块名称,
                    p.地块编码,
                    p.水田,
                    p.旱地,
                    "",
                    p.面积,
                    "",
                    p.东至,
                    p.南至,
                    p.西至,
                    p.北至
                };

                //若超过数量新增表单
                if (groupOther.Count() > 18 || groupZ.Count() > 4)
                {
                    ISheet newSheet = 登记表.Workbook.CloneSheet(2);

                    var row4 = new Excel.Export.AutoCode.登记表.Row(newSheet.SheetName, 4, 3);
                    var row26 = new Excel.Export.AutoCode.登记表.Row(newSheet.SheetName, 23, 3);

                    row4.Fill(groupOther.Skip(18).Take(18), func);

                    row26.Fill(groupZ.Skip(4).Take(4), func);
                }

                //若超过数量新增表单...
                if (groupOther.Count() > 36 || groupZ.Count() > 8)
                {
                    ISheet newSheet = 登记表.Workbook.CloneSheet(2);

                    var row4 = new Excel.Export.AutoCode.登记表.Row(newSheet.SheetName, 4, 3);
                    var row26 = new Excel.Export.AutoCode.登记表.Row(newSheet.SheetName, 23, 3);

                    row4.Fill(groupOther.Skip(36).Take(18), func);

                    row26.Fill(groupZ.Skip(8).Take(4), func);
                }

                if (groupOther.Count() > 54 || groupZ.Count() > 12)
                {
                    Console.WriteLine("承包方代表编码为：{0} 的承包地或自留地太多， 未处理完。。。", 承包方.承包方代表编码);
                }

                登记表.登记簿2.Row4a.Fill(二轮承包信息, p => new List<Object>
                {
                    p.地块名称,
                    p.面积
                });

                登记表.登记簿2.Row4.Fill(groupOther.Take(18), func);

                登记表.登记簿2.Row23.Fill(groupZ.Take(4), func);

                ////计算公式
                //登记表.Workbook.GetSheetAt(1).ForceFormulaRecalculation = true;
                //登记表.Workbook.GetSheetAt(2).ForceFormulaRecalculation = true;

                //导出文件
                var newFile = File.Create(String.Format(file, first地块.承包方代表编码));
                登记表.Workbook.Write(newFile);
                newFile.Close();
            }
        }

        public static Object ToNumber(Boolean boolValue, Object dataValue)
        {
            if (boolValue)
                return dataValue;
            else
                return null;
        }
    }
}
