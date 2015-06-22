using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Model {
public class 登记表
{
public static IWorkbook Workbook { get; set; }
public class 封面
{
private const String sheetName = "封面";
public static Cell 行政区名称 = new Cell(sheetName, 2, 1);
public static Cell 承包方代表姓名 = new Cell(sheetName, 2, 2);
public static Cell 个体编码 = new Cell(sheetName, 2, 3);
}

public class 登记簿1
{
private const String sheetName = "登记簿1";
public static Cell 个体编码 = new Cell(sheetName, 2, 1);
public static Cell 行政区名称 = new Cell(sheetName, 2, 2);
public static Cell 承包方代表姓名 = new Cell(sheetName, 2, 3);
public static Cell 邮编 = new Cell(sheetName, 7, 3);
public static Cell 证件号码 = new Cell(sheetName, 2, 4);
public static Cell 非在册户口 = new Cell(sheetName, 8, 4);
public static Cell 承包方地址 = new Cell(sheetName, 2, 5);
public static Cell 个体编码1 = new Cell(sheetName, 2, 6);
public static Row  Row11 = new Row(sheetName, 11, 0);
public static Row  Row26 = new Row(sheetName, 26, 0);
}

public class 登记簿2
{
private const String sheetName = "登记簿2";
public static Row  Row4 = new Row(sheetName, 4, 3);
public static Row  Row23 = new Row(sheetName, 23, 3);
}

public class 登记簿3
{
private const String sheetName = "登记簿3";
}

public class 登记簿四附农户地块示意图
{
private const String sheetName = "登记簿（四）附农户地块示意图";
}

 public class Cell
        {
            private String SheetName { get; set; }
            private Int32 X { get; set; }
            private Int32 Y { get; set; }
            public Cell(String sheetName, Int32 x, Int32 y)
            {
                SheetName = sheetName; X = x; Y = y;
            }
            public void Fill(Object value)
            {
                SetCellValue(Workbook.GetSheet(SheetName).GetRow(Y).GetCell(X), value);
            }
        }

    public class Row
        {
            private String SheetName { get; set; }
            private Int32 BeginRow { get; set; }
            private Int32 BeginColum { get; set; }
            public Row(String sheetName, Int32 beginRow, Int32 beginColum)
            {
                SheetName = sheetName;
                BeginRow = beginRow;
                BeginColum = beginColum;
            }
            public void Fill<T>(IEnumerable<T> dataList, Func<T, List<Object>> setAction)
            {
                Int32 beginIndex = BeginRow;

                foreach (var data in dataList)
                {
                    var itemVaues = setAction(data);

                    for (int j = BeginColum; j < itemVaues.Count + BeginColum; j++)
                    {
                        var cell = Workbook.GetSheet(SheetName).GetRow(beginIndex).GetCell(j);

                        SetCellValue(cell, itemVaues[j - BeginColum]);
                    }
                    beginIndex++;
                }
            }
        }

            public static void SetCellValue(ICell cell, Object value)
            {
                if (cell == null) return;
                if (value == null)
                {
                    cell.SetCellValue(String.Empty);
                    return;
                }

                var valueTypeCode = Type.GetTypeCode(value.GetType());
                switch (valueTypeCode)
                {
                    case TypeCode.String:
                        cell.SetCellValue(Convert.ToString(value)); break;
                    case TypeCode.DateTime:
                        cell.SetCellValue(Convert.ToDateTime(value)); break;
                    case TypeCode.Boolean:
                        cell.SetCellValue(Convert.ToBoolean(value)); break;
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.Byte:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                        cell.SetCellValue(Convert.ToDouble(value)); break;
                    default:
                        cell.SetCellValue(string.Empty); break;
                }
            }

}
}
