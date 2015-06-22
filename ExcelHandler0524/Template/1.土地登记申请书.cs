using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace Model {
public class 1.土地登记申请书
{
public static IWorkbook Workbook { get; set; }
public class Sheet1
{
private const String sheetName = "Sheet1";
public static Cell 宗地代码 = new Cell(sheetName, 1, 1);
public static Cell 土地权利人 = new Cell(sheetName, 3, 6);
public static Cell 证件编号 = new Cell(sheetName, 3, 8);
public static Cell 通讯地址 = new Cell(sheetName, 3, 10);
public static Cell 土地坐落 = new Cell(sheetName, 3, 26);
public static Cell 宗地面积 = new Cell(sheetName, 3, 27);
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
