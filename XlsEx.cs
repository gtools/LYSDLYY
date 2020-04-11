using System;
using Spire.Xls;
using System.Data;
using System.Drawing;

namespace LYSDLYY
{
    /// <summary>
    /// xls扩展类
    /// </summary>
    public static class XlsEx
    {
        //合并单元格
        //workbook.Worksheets[0].Range["A3:B5"].Merge();
        //取消合并单元格
        //workbook.Worksheets[0].Range["A3:B5"].UnMerge();
        #region 单元格
        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static CellRange GetCell(this Worksheet sheet, int row, int column)
        {
            return sheet.Range[row, column];
        }
        /// <summary>
        /// 获取单元格,1,1,
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static CellRange GetCellFirst(this Worksheet sheet)
        {
            return sheet.Range[1, 1];
        }
        /// <summary>
        /// 获取单元格集合列
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static CellRange GetColumn(this Worksheet sheet, int column)
        {
            return sheet.Columns[column];
        }
        /// <summary>
        /// 获取单元格集合行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        public static CellRange GetRow(this Worksheet sheet, int row)
        {
            // sheet.Rows[row]从0开始，sheet.Range[row, column]从1开始;
            return sheet.Rows[row - 1];
        }

        /// <summary>
        /// 获取单元格集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="lastRow"></param>
        /// <param name="lastColumn"></param>
        /// <returns></returns>
        public static CellRange GetCell(this Worksheet sheet, int row, int column, int lastRow, int lastColumn)
        {
            return sheet.Range[row, column, lastRow, lastColumn];
        }
        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string SetCell(this CellRange cell, string value)
        {
            return cell.Text = value;
        }
        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static double SetCell(this CellRange cell, double value)
        {
            return cell.NumberValue = value;
        }
        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DateTime SetCell(this CellRange cell, DateTime value)
        {
            return cell.DateTimeValue = value;
        }
        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool SetCell(this CellRange cell, bool value)
        {
            return cell.BooleanValue = value;
        }
        /// <summary>
        /// 替换单元格数据
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        public static string SetCellReplace(this CellRange cell, string oldValue, string newValue)
        {
            return cell.SetCell(cell.Text.Replace(oldValue, newValue));
        }
        #endregion


        /// <summary>
        /// 数据导入excel
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        /// <param name="firstRow"></param>
        public static void DataTableToExcel(this Worksheet sheet, DataTable dt, int firstRow)
        {
            sheet.InsertDataTable(dt, false, firstRow, 1);
        }
        /// <summary>
        /// 数据导入excel,第一行第一列
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dt"></param>
        public static void DataTableToExcel(this Worksheet sheet, DataTable dt)
        {
            sheet.DataTableToExcel(dt, 1);
        }
        /// <summary>
        /// 添加边框
        /// </summary>
        /// <param name="cells"></param>
        public static CellRange StyleLine(this CellRange cells)
        {
            //cells.Style.Borders.LineStyle = LineStyleType.Medium;
            cells.BorderInside(LineStyleType.Thin);
            cells.BorderAround(LineStyleType.Thin);
            cells.Borders.KnownColor = ExcelColors.Black;
            return cells;
        }
        /// <summary>
        /// 添加字体为红色
        /// </summary>
        /// <param name="cells"></param>
        public static CellRange StyleFontColorRed(this CellRange cells)
        {
            cells.StyleFontColor(Color.Red);
            return cells;
        }
        /// <summary>
        /// 添加字体颜色
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="color"></param>
        /// <returns></returns>
        public static CellRange StyleFontColor(this CellRange cells, Color color)
        {
            cells.Style.Font.Color = color;
            return cells;
        }
        /// <summary>
        /// 添加字体大小
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="size"></param>
        /// <returns></returns>
        public static CellRange StyleFontSize(this CellRange cells, double size)
        {
            cells.Style.Font.FontName = "宋体";
            cells.Style.Font.Size = size;
            return cells;
        }
        /// <summary>
        /// 添加字体粗体
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="isbold"></param>
        /// <returns></returns>
        public static CellRange StyleFontIsBold(this CellRange cells, bool isbold)
        {
            cells.Style.Font.IsBold = isbold;
            return cells;
        }
        /// <summary>
        /// 添加字体粗体
        /// </summary>
        /// <param name="cells"></param>
        /// <returns></returns>
        public static CellRange StyleFontIsBold(this CellRange cells)
        {
            cells.StyleFontIsBold(true);
            return cells;
        }
        /// <summary>
        /// 添加字体居中
        /// </summary>
        /// <param name="cells"></param>
        /// <returns></returns>
        public static CellRange StyleFontCenter(this CellRange cells)
        {
            cells.HorizontalAlignment = HorizontalAlignType.Center;
            cells.VerticalAlignment = VerticalAlignType.Center;
            return cells;
        }
    }
}