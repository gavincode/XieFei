using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel.Export.AutoCode;
using NPOI.SS.UserModel;

namespace CodeGenerator
{
    public class TemplateHandler
    {
        const String CellPrefix = "#";
        const String TableCellPrefix = "&";
        static Regex strRegex = new Regex(String.Format(@"(?<={0}\b*?)([\w]+)", CellPrefix));
        static Regex strTableRegex = new Regex(String.Format(@"(?<={0})([\w]+)", TableCellPrefix));
        static Regex numRegex = new Regex(@"(?<=[\w]+)([0-9]+)");

        public static void Create(String file)
        {
            登记表.Workbook = WorkbookFactory.Create(file);

            var placeHolderInfo = GetPlaceHolderInfo(登记表.Workbook);
            var placeTableInfo = GetRowPlaceHolderInfo(登记表.Workbook);

            ////
            //登记表.封面.承包方代表姓名.Fill("承包方代表姓名");
            //登记表.封面.个体编码.Fill("个体编码");
            //登记表.封面.行政区名称.Fill("行政区名称");

            //登记表.登记簿1.承包方代表姓名.Fill("aaa");
            //登记表.登记簿1.邮编.Fill(1000);


            //登记表.Workbook.Write(File.Create("test.xls"));

            StringBuilder codeBuilder = new StringBuilder();

            codeBuilder.AppendLine("using System;");
            codeBuilder.AppendLine("using System.Collections.Generic;");
            codeBuilder.AppendLine("using System.Linq;");
            codeBuilder.AppendLine("using System.Text;");
            codeBuilder.AppendLine("using NPOI.SS.UserModel;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("namespace Model {");
            codeBuilder.AppendLine(String.Format("public class {0}", TrimFileName(Path.GetFileNameWithoutExtension(file))));
            codeBuilder.AppendLine("{");
            codeBuilder.AppendLine("public static IWorkbook Workbook { get; set; }");

            foreach (var sheetInfo in placeHolderInfo)
            {
                String className = RemoveDisableSymbols(sheetInfo.Key);
                StringBuilder sheetBuilder = new StringBuilder();
                sheetBuilder.AppendLine(String.Format("public class {0}", className));
                sheetBuilder.AppendLine("{");
                sheetBuilder.AppendLine(String.Format("private const String sheetName = \"{0}\";", sheetInfo.Key));

                foreach (var para in sheetInfo.Value)
                {
                    sheetBuilder.AppendLine(String.Format("public static Cell {2} = new Cell(sheetName, {0}, {1});", para.Value.X, para.Value.Y, para.Key));
                }

                foreach (var group in placeTableInfo[sheetInfo.Key].GroupBy(p => p.Y))
                {
                    Int32 beginColum = group.Min(p => p.X);
                    sheetBuilder.AppendLine(String.Format("public static Row  Row{0} = new Row(sheetName, {0}, {1});", group.Key, beginColum));
                }

                sheetBuilder.AppendLine("}");

                codeBuilder.AppendLine(sheetBuilder.ToString());
            }

            codeBuilder.AppendLine(GetCellClass());
            codeBuilder.AppendLine(GetRowClass());
            codeBuilder.AppendLine(GetMethod());
            codeBuilder.AppendLine("}");
            codeBuilder.AppendLine("}");

            var code = codeBuilder.ToString();

            var stream = File.CreateText(Path.ChangeExtension(file, ".cs"));
            stream.Write(code);
            stream.Close();
        }

        private static Dictionary<String, Dictionary<String, Point>> GetPlaceHolderInfo(IWorkbook workbook)
        {
            Dictionary<String, Dictionary<String, Point>> placeHolderInfo = new Dictionary<String, Dictionary<String, Point>>();

            foreach (ISheet sheet in workbook)
            {
                Dictionary<String, Point> sheetCellInfos = new Dictionary<String, Point>();
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType.Equals(CellType.String))
                        {
                            MatchCollection matches = strRegex.Matches(cell.StringCellValue);
                            foreach (Match match in matches)
                            {
                                AddIn(sheetCellInfos, match.Value, new Point(cell.ColumnIndex, cell.RowIndex));
                            }
                        }
                    }
                }

                placeHolderInfo[sheet.SheetName] = sheetCellInfos;
            }

            return placeHolderInfo;
        }

        private static Dictionary<String, List<Point>> GetRowPlaceHolderInfo(IWorkbook workbook)
        {
            Dictionary<String, List<Point>> placeHolderInfo = new Dictionary<String, List<Point>>();

            foreach (ISheet sheet in workbook)
            {
                List<Point> sheetCellInfos = new List<Point>();
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        if (cell.CellType.Equals(CellType.String))
                        {
                            MatchCollection matches = strTableRegex.Matches(cell.StringCellValue);
                            foreach (Match match in matches)
                            {
                                sheetCellInfos.Add(new Point(cell.ColumnIndex, cell.RowIndex));
                            }
                        }
                    }
                }

                placeHolderInfo[sheet.SheetName] = sheetCellInfos;
            }

            return placeHolderInfo;
        }

        private static void AddIn(Dictionary<String, Point> dict, String key, Point value)
        {
            if (!dict.ContainsKey(key))
            {
                dict[key] = value;
                return;
            }

            var matches = numRegex.Matches(key);

            if (matches.Count == 0)
                AddIn(dict, key + "1", value);
            else
                AddIn(dict, key.Replace(matches[0].Value, (Convert.ToInt32(matches[0].Value) + 1).ToString()), value);
        }

        private static String RemoveDisableSymbols(String src)
        {
            const String symbols = "(|)|（|）";
            foreach (var item in symbols.Split('|'))
            {
                if (src.Contains(item))
                    src = src.Replace(item, String.Empty);
            }

            return src;
        }

        private static String GetCellClass()
        {
            StringBuilder codeBuilder = new StringBuilder();
            codeBuilder.AppendLine(" public class Cell");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            private String SheetName { get; set; }");
            codeBuilder.AppendLine("            private Int32 X { get; set; }");
            codeBuilder.AppendLine("            private Int32 Y { get; set; }");
            codeBuilder.AppendLine("            public Cell(String sheetName, Int32 x, Int32 y)");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                SheetName = sheetName; X = x; Y = y;");
            codeBuilder.AppendLine("            }");
            codeBuilder.AppendLine("            public void Fill(Object value)");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                SetCellValue(Workbook.GetSheet(SheetName).GetRow(Y).GetCell(X), value);");
            codeBuilder.AppendLine("            }");
            codeBuilder.AppendLine("            public String GetValue()");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                return Workbook.GetSheet(SheetName).GetRow(Y).GetCell(X).StringCellValue;");
            codeBuilder.AppendLine("            }");
            codeBuilder.AppendLine("        }");

            return codeBuilder.ToString();
        }

        private static String GetRowClass()
        {
            StringBuilder codeBuilder = new StringBuilder();
            codeBuilder.AppendLine("    public class Row");
            codeBuilder.AppendLine("        {");
            codeBuilder.AppendLine("            private String SheetName { get; set; }");
            codeBuilder.AppendLine("            private Int32 BeginRow { get; set; }");
            codeBuilder.AppendLine("            private Int32 BeginColum { get; set; }");
            codeBuilder.AppendLine("            public Row(String sheetName, Int32 beginRow, Int32 beginColum)");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                SheetName = sheetName;");
            codeBuilder.AppendLine("                BeginRow = beginRow;");
            codeBuilder.AppendLine("                BeginColum = beginColum;");
            codeBuilder.AppendLine("            }");
            codeBuilder.AppendLine("            public void Fill<T>(IEnumerable<T> dataList, Func<T, List<Object>> setAction)");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                Int32 beginIndex = BeginRow;");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("                foreach (var data in dataList)");
            codeBuilder.AppendLine("                {");
            codeBuilder.AppendLine("                    var itemVaues = setAction(data);");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("                    for (int j = BeginColum; j < itemVaues.Count + BeginColum; j++)");
            codeBuilder.AppendLine("                    {");
            codeBuilder.AppendLine("                        var cell = Workbook.GetSheet(SheetName).GetRow(beginIndex).GetCell(j);");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("                        SetCellValue(cell, itemVaues[j - BeginColum]);");
            codeBuilder.AppendLine("                    }");
            codeBuilder.AppendLine("                    beginIndex++;");
            codeBuilder.AppendLine("                }");
            codeBuilder.AppendLine("            }");
            codeBuilder.AppendLine("        }");

            return codeBuilder.ToString();
        }

        private static String GetMethod()
        {
            StringBuilder codeBuilder = new StringBuilder();

            codeBuilder.AppendLine("            public static void SetCellValue(ICell cell, Object value)");
            codeBuilder.AppendLine("            {");
            codeBuilder.AppendLine("                if (cell == null) return;");
            codeBuilder.AppendLine("                if (value == null)");
            codeBuilder.AppendLine("                {");
            codeBuilder.AppendLine("                    cell.SetCellValue(String.Empty);");
            codeBuilder.AppendLine("                    return;");
            codeBuilder.AppendLine("                }");
            codeBuilder.AppendLine("");
            codeBuilder.AppendLine("                var valueTypeCode = Type.GetTypeCode(value.GetType());");
            codeBuilder.AppendLine("                switch (valueTypeCode)");
            codeBuilder.AppendLine("                {");
            codeBuilder.AppendLine("                    case TypeCode.String:");
            codeBuilder.AppendLine("                        cell.SetCellValue(Convert.ToString(value)); break;");
            codeBuilder.AppendLine("                    case TypeCode.DateTime:");
            codeBuilder.AppendLine("                        cell.SetCellValue(Convert.ToDateTime(value)); break;");
            codeBuilder.AppendLine("                    case TypeCode.Boolean:");
            codeBuilder.AppendLine("                        cell.SetCellValue(Convert.ToBoolean(value)); break;");
            codeBuilder.AppendLine("                    case TypeCode.Int16:");
            codeBuilder.AppendLine("                    case TypeCode.Int32:");
            codeBuilder.AppendLine("                    case TypeCode.Int64:");
            codeBuilder.AppendLine("                    case TypeCode.Byte:");
            codeBuilder.AppendLine("                    case TypeCode.Single:");
            codeBuilder.AppendLine("                    case TypeCode.Double:");
            codeBuilder.AppendLine("                    case TypeCode.UInt16:");
            codeBuilder.AppendLine("                    case TypeCode.UInt32:");
            codeBuilder.AppendLine("                    case TypeCode.UInt64:");
            codeBuilder.AppendLine("                        cell.SetCellValue(Convert.ToDouble(value)); break;");
            codeBuilder.AppendLine("                    default:");
            codeBuilder.AppendLine("                        cell.SetCellValue(string.Empty); break;");
            codeBuilder.AppendLine("                }");
            codeBuilder.AppendLine("            }");

            return codeBuilder.ToString();
        }

        private static String TrimFileName(String fileName)
        {
            var array = fileName.Split('.');

            return array.Length == 2 ? array[1] : fileName;
        }

        public class Point
        {
            public int X { get; set; }
            public int Y { get; set; }

            public Point(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }
}
