using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using BLL;
using Model;
using NPOI.SS.UserModel;

namespace ExcelHandler
{
    public class ExcelHandler0524
    {
        const String templateFile1 = @"Template\1.土地登记申请书.xls";
        const String templateFile2 = @"Template\2.地籍调查表.xls";
        const String templateFile3 = @"Template\3.土地登记审批表.xls";
        const String templateFile4 = @"Template\4.土地登记卡.xls";
        const String templateFile5 = @"Template\5.归户卡.xls";

        const String file = @"Data\{0}.xls";

        public static void Init()
        {
            if (!Directory.Exists("Data"))
                Directory.CreateDirectory("Data");

            if (!File.Exists(templateFile1)
                || !File.Exists(templateFile2)
                || !File.Exists(templateFile3)
                || !File.Exists(templateFile4)
                || !File.Exists(templateFile5)
                )
            {
                Console.WriteLine("模板不存在!");
            }
        }

        public static void ExportTemplate()
        {
            var 宗地属性 = CommonBLL.宗地属性();

            foreach (var model in 宗地属性)
            {
                土地登记申请书.Workbook = WorkbookFactory.Create(templateFile1);
                地籍调查表.Workbook = WorkbookFactory.Create(templateFile2);
                土地登记审批表.Workbook = WorkbookFactory.Create(templateFile3);
                土地登记卡.Workbook = WorkbookFactory.Create(templateFile4);
                归户卡.Workbook = WorkbookFactory.Create(templateFile5);

                //土地登记申请书
                土地登记申请书.Sheet1.通讯地址.Fill(model.通讯地址);
                土地登记申请书.Sheet1.土地权利人.Fill(model.土地权利人);
                土地登记申请书.Sheet1.证件编号.Fill(model.证件编号);
                土地登记申请书.Sheet1.宗地代码.Fill(model.宗地代码);
                土地登记申请书.Sheet1.土地坐落.Fill(model.土地坐落);
                土地登记申请书.Sheet1.原证面积.Fill(model.原证面积);

                //地籍调查表
                地籍调查表.Sheet1.土地权利人.Fill(model.土地权利人);
                地籍调查表.Sheet1.宗地代码1.Fill(model.宗地代码);
                地籍调查表.Sheet1.宗地代码.Fill(model.宗地代码);

                地籍调查表.Sheet2.北至.Fill("北至: " + model.北至);
                地籍调查表.Sheet2.东至.Fill("东至: " + model.东至);
                地籍调查表.Sheet2.南至.Fill("南至: " + model.南至);
                地籍调查表.Sheet2.西至.Fill("西至: " + model.西至);
                地籍调查表.Sheet2.批准面积.Fill(model.批准面积);
                地籍调查表.Sheet2.通讯地址.Fill(model.通讯地址);
                地籍调查表.Sheet2.图幅号.Fill(model.图幅号);
                地籍调查表.Sheet2.土地权利人.Fill(model.土地权利人);
                地籍调查表.Sheet2.土地坐落.Fill(model.土地坐落);
                地籍调查表.Sheet2.证件编号.Fill(model.证件编号);
                地籍调查表.Sheet2.宗地代码.Fill(model.宗地代码);
                地籍调查表.Sheet2.宗地面积.Fill(model.宗地面积);

                地籍调查表.Sheet4.权利人名称.Fill(model.土地权利人);
                地籍调查表.Sheet4.土地坐落.Fill(model.土地坐落);
                地籍调查表.Sheet4.宗地代码.Fill(model.宗地代码);
                地籍调查表.Sheet4.宗地面积.Fill(model.宗地面积);

                //土地登记审批表
                土地登记审批表.Sheet1.宗地代码.Fill(model.宗地代码);

                土地登记审批表.Sheet2.土地权利人.Fill(model.土地权利人);
                土地登记审批表.Sheet2.证件编号.Fill(model.证件编号);
                土地登记审批表.Sheet2.通讯地址.Fill(model.通讯地址);
                土地登记审批表.Sheet2.宗地代码.Fill(model.宗地代码);
                土地登记审批表.Sheet2.图幅号.Fill(model.图幅号);
                土地登记审批表.Sheet2.土地坐落.Fill(model.土地坐落);
                土地登记审批表.Sheet2.宗地面积.Fill(model.宗地面积);
                土地登记审批表.Sheet2.批准面积.Fill(model.批准面积);
                土地登记审批表.Sheet2.宗地代码1.Fill(model.宗地代码);
                土地登记审批表.Sheet2.原证书号.Fill(model.原证书号);
                土地登记审批表.Sheet2.核实面积文本.Fill(model.核实面积文本);

                //土地登记卡
                土地登记卡.Sheet1.宗地代码.Fill(model.宗地代码);
                土地登记卡.Sheet1.图幅号.Fill(model.图幅号);
                土地登记卡.Sheet1.宗地面积.Fill(model.宗地面积);
                土地登记卡.Sheet1.土地坐落.Fill(model.土地坐落);
                土地登记卡.Sheet1.土地权利人.Fill(model.土地权利人);
                土地登记卡.Sheet1.通讯地址.Fill(model.通讯地址);
                土地登记卡.Sheet1.批准面积.Fill(model.批准面积);
                土地登记卡.Sheet1.证件编号.Fill(model.证件编号);
                土地登记卡.Sheet1.新证书号.Fill(model.新证书号);

                //归户卡
                归户卡.Sheet1.土地权利人.Fill(model.土地权利人);
                归户卡.Sheet1.通讯地址.Fill(model.通讯地址);
                归户卡.Sheet1.证件编号.Fill(model.证件编号);
                归户卡.Sheet1.宗地代码.Fill(model.宗地代码);

                var query = 宗地属性.Where(p => p.证件编号 == model.证件编号);

                Func<宗地属性, List<Object>> func = p => new List<Object> 
                { 
                    p.宗地代码,
                    p.图幅号,
                    p.新证书号,
                    p.新证书号,
                    p.土地坐落,
                    p.土地坐落,
                    "宅基地使用权",
                    "批准拨用宅基地",
                    "农村宅基地",
                    p.批准面积
                };

                if (model.证件编号 != null && model.证件编号 != "#N/A")
                    归户卡.Sheet1.Row5.Fill(query, func);

                var dir = Directory.CreateDirectory(String.Format("Data\\{0}-{1}", model.土地权利人, model.宗地代码));
                var file1 = File.Create(Path.Combine(dir.FullName, "1.土地登记申请书.xls"));
                var file2 = File.Create(Path.Combine(dir.FullName, "2.地籍调查表.xls"));
                var file3 = File.Create(Path.Combine(dir.FullName, "3.土地登记审批表.xls"));
                var file4 = File.Create(Path.Combine(dir.FullName, "4.土地登记卡.xls"));
                var file5 = File.Create(Path.Combine(dir.FullName, "5.归户卡.xls"));

                土地登记申请书.Workbook.Write(file1);
                地籍调查表.Workbook.Write(file2);
                土地登记审批表.Workbook.Write(file3);
                土地登记卡.Workbook.Write(file4);
                归户卡.Workbook.Write(file5);

                file1.Close();
                file2.Close();
                file3.Close();
                file4.Close();
                file5.Close();
            }
        }
    }
}