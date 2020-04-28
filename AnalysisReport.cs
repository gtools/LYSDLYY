using GTSharp;
using Spire.Xls;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using GTSharp.Extension;
using GTSharp.IO;

namespace LYSDLYY
{
    /// <summary>
    /// 分析统计
    /// </summary>
    public class AnalysisReport
    {
        #region 每周院长查询报表
        /// <summary>
        /// 模板：每周院长查询报表
        /// 导出：每周院长查询报表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MZYZCXBB1(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy().AsEnumerable().Take(7).CopyToDataTable();
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月"));
            // 设置单元格计数
            sheet.GetCellFirst().SetCellReplace("[NUM]", (Helper.GetWeekNumInMonth(Date) - 1).ToString());
            sheet.DataTableToExcel(Data, RowBeginIndex);

            book.SaveToFile(Path.Combine(PathSave, NameSave));
            sheet.Dispose();
            book.Dispose();

            //初始化workbook实例
            Workbook workbook = new Workbook();
            //加载Excel文档
            workbook.LoadFromFile(Path.Combine(PathSave, NameSave));
            // 获取第一个工作表
            Worksheet sheet1 = workbook.Worksheets[0];
            //将图表保存为图片
            Image[] imgs = workbook.SaveChartAsImage(sheet1);
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, "{0}" + NameSave), "png");
            DirectoryHelper.Create(Path.GetDirectoryName(PathSaveImage));
            for (int i = 0; i < imgs.Length; i++)
            {
                imgs[i].Save(string.Format(PathSaveImage, i + 1), ImageFormat.Png);
            }
            sheet1.SaveToImage(1, 1, RowEndIndex + 1, Data.Columns.Count).Save(string.Format(PathSaveImage, string.Empty), ImageFormat.Png);
            // 处理白边
            Bitmap bitmap = new Bitmap(string.Format(PathSaveImage, string.Empty));
            Bitmap bitmap1 = Helper.KiCut(bitmap, 66, 66, bitmap.Width - 66 - 66, bitmap.Height - 66 - 66);
            bitmap.Dispose();
            bitmap1.Save(string.Format(PathSaveImage, string.Empty));
        }
        #endregion

        #region 每日院长查询报表
        /// <summary>
        /// 模板：每日1科室在院人数一览表
        /// 导出：科室在院人数一览表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB1(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 添加边框
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).StyleLine();
            // 添加字体红色
            sheet.GetRow(RowEndIndex).StyleFontColorRed();
            // 合计为0删除列
            for (int i = Data.Columns.Count; i >= 1; i--)
            {
                //获取单元格数据
                var _temp = sheet.GetCell(RowEndIndex, i).Text;
                //删除列
                if (_temp == "0" || _temp.IsNullOrWhiteSpace())
                    sheet.DeleteColumn(i);
            }
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每日2按手术时间统计手术人数表
        /// 导出：按手术时间统计手术人数表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB2(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            // 设置单元格计数
            sheet.GetCell(2, 1).SetCellReplace("[NUM]", Data.Rows.Count.ToString());
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 添加边框 // 字号 // 加粗 // 居中
            var cells = sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count);
            cells.StyleLine().StyleFontSize(16).StyleFontIsBold().StyleFontCenter();
            // 行高
            cells.RowHeight = 32;
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每日3在院危重病人患者明细表
        /// 导出：在院危重病人患者明细表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB3(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            // 设置单元格计数
            sheet.GetCell(2, 1).SetCellReplace("[NUM]", Data.Rows.Count.ToString());
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 添加边框 // 字号 // 加粗 // 居中
            var cells = sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count);
            cells.StyleLine().StyleFontSize(16).StyleFontIsBold().StyleFontCenter();
            // 行高
            cells.RowHeight = 32;
            // 判断日期字体变红
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                var _temp = sheet.GetCell(i + RowBeginIndex, Data.Columns.Count).Text;
                if (_temp == Date.ToString("yyyy-MM-dd") || _temp == Date.AddDays(-1).ToString("yyyy-MM-dd"))
                    sheet.GetRow(i + RowBeginIndex).StyleFontColorRed();
            }
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /*
        /// <summary>
        /// 模板：每日4在院I级护理患者明细表
        /// 导出：在院I级护理患者明细表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB4(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            // 设置单元格计数
            sheet.GetCell(2, 1).SetCellReplace("[NUM]", Data.Rows.Count.ToString());
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 添加边框 // 字号 // 加粗 // 居中
            var cells = sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count);
            cells.StyleLine().StyleFontSize(16).StyleFontIsBold().StyleFontCenter();
            // 行高
            cells.RowHeight = 32;
            //判断日期字体变红
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                var _temp = sheet.GetCell(i + RowBeginIndex, Data.Columns.Count).Text;
                if (_temp == Date.ToString("yyyy-MM-dd") || _temp == Date.AddDays(-1).ToString("yyyy-MM-dd"))
                    sheet.GetRow(i + RowBeginIndex).StyleFontColorRed();
            }
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            DirectoryHelper.Create(Path.GetDirectoryName(PathSaveImage));
            sheet.SaveToImage(PathSaveImage, ImageFormat.Png);
            // 处理白边
            Bitmap bitmap = new Bitmap(PathSaveImage);
            Bitmap bitmap1 = Helper.KiCut(bitmap, 66, 66, bitmap.Width - 66 - 66, bitmap.Height - 66 - 66);
            bitmap.Dispose();
            bitmap1.Save(PathSaveImage);
        }
        */
        /// <summary>
        /// 模板：每日5主要业务数据表
        /// 导出：主要业务数据表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB5(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            //设置单元格数据
            //全院收入
            sheet.SetCellValue(5, 6, Data.Rows[0][0].ToString());
            //住院收入
            sheet.SetCellValue(4, 6, Data.Rows[0][1].ToString());
            //门诊收入
            sheet.SetCellValue(3, 6, Data.Rows[0][2].ToString());
            //全院药品收入
            sheet.SetCellValue(5, 2, Data.Rows[0][3].ToString());
            //住院药品收入
            sheet.SetCellValue(4, 2, Data.Rows[0][4].ToString());
            //门诊药品收入
            sheet.SetCellValue(3, 2, Data.Rows[0][5].ToString());
            //全院药占比
            sheet.SetCellValue(5, 3, Data.Rows[0][6].ToString());
            //住院药占比
            sheet.SetCellValue(4, 3, Data.Rows[0][7].ToString());
            //门诊药占比
            sheet.SetCellValue(3, 3, Data.Rows[0][8].ToString());
            //全院人次
            //sheet.SetCellValue(_dtstartheight, 2, Data.Rows[0][9].ToString());
            //住院人次
            sheet.SetCellValue(4, 4, Data.Rows[0][10].ToString());
            //门诊人次
            sheet.SetCellValue(3, 4, Data.Rows[0][11].ToString());
            //全院平均
            //sheet.SetCellValue(_dtstartheight, 2, _dt.Rows[0][12].ToString());
            //住院平均
            sheet.SetCellValue(4, 5, Data.Rows[0][13].ToString());
            //门诊平均
            sheet.SetCellValue(3, 5, Data.Rows[0][14].ToString());
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每日8主要业务数据表
        /// 导出：主要业务数据表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MRYYCXBB8(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            //设置单元格数据
            //全院收入
            sheet.SetCellValue(5, 6, Data.Rows[0][0].ToString());
            //住院收入
            sheet.SetCellValue(4, 6, Data.Rows[0][1].ToString());
            //门诊收入
            sheet.SetCellValue(3, 6, Data.Rows[0][2].ToString());
            //全院药品收入
            sheet.SetCellValue(5, 2, Data.Rows[0][3].ToString());
            //住院药品收入
            sheet.SetCellValue(4, 2, Data.Rows[0][4].ToString());
            //门诊药品收入
            sheet.SetCellValue(3, 2, Data.Rows[0][5].ToString());
            //全院药占比
            sheet.SetCellValue(5, 3, Data.Rows[0][6].ToString());
            //住院药占比
            sheet.SetCellValue(4, 3, Data.Rows[0][7].ToString());
            //门诊药占比
            sheet.SetCellValue(3, 3, Data.Rows[0][8].ToString());
            //全院人次
            //sheet.SetCellValue(_dtstartheight, 2, Data.Rows[0][9].ToString());
            //住院人次
            sheet.SetCellValue(4, 4, Data.Rows[0][10].ToString());
            //门诊人次
            sheet.SetCellValue(3, 4, Data.Rows[0][11].ToString());
            //全院平均
            //sheet.SetCellValue(_dtstartheight, 2, _dt.Rows[0][12].ToString());
            //住院平均
            sheet.SetCellValue(4, 5, Data.Rows[0][13].ToString());
            //门诊平均
            sheet.SetCellValue(3, 5, Data.Rows[0][14].ToString());
            //去年当天在院人数
            sheet.SetCellValue(7, 2, Data.Rows[0][15].ToString());
            //去年当天出院人次
            sheet.SetCellValue(8, 2, Data.Rows[0][16].ToString());
            //去年当天住院总收入
            sheet.SetCellValue(9, 2, Data.Rows[0][17].ToString());
            //去年当天门诊人次
            sheet.SetCellValue(10, 2, Data.Rows[0][18].ToString());
            //去年当天门诊总收入
            sheet.SetCellValue(11, 2, Data.Rows[0][19].ToString());
            //去年当天全院总收入
            sheet.SetCellValue(12, 2, Data.Rows[0][20].ToString());
            //在院人数同比
            sheet.SetCellValue(7, 4, Data.Rows[0][21].ToString());
            //出院结算人次同比
            sheet.SetCellValue(8, 4, Data.Rows[0][22].ToString());
            //住院总收入同比
            sheet.SetCellValue(9, 4, Data.Rows[0][23].ToString());
            //门诊结算人次同比
            sheet.SetCellValue(10, 4, Data.Rows[0][24].ToString());
            //门诊总收入同比
            sheet.SetCellValue(11, 4, Data.Rows[0][25].ToString());
            //全院总收入同比
            sheet.SetCellValue(12, 4, Data.Rows[0][26].ToString());
            //在院人数
            sheet.SetCellValue(7, 6, Data.Rows[0][27].ToString());
            //入院人次
            sheet.SetCellValue(8, 6, Data.Rows[0][28].ToString());
            //急诊人次
            sheet.SetCellValue(9, 6, Data.Rows[0][29].ToString());
            //手术台次
            sheet.SetCellValue(10, 6, Data.Rows[0][30].ToString());
            //危重患者人数
            sheet.SetCellValue(11, 6, Data.Rows[0][31].ToString());
            //一级护理人数
            sheet.SetCellValue(12, 6, Data.Rows[0][32].ToString());
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        #endregion

        #region 每月院长查询报表
        /// <summary>
        /// 模板：每月1住院主要业务数据同期比表
        /// 导出：住院主要业务数据同期比表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MYYZCXBB1(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '判断
            var BMonth = bool.Parse(com.GetParam(9));
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            if (BMonth)
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date:yyyy年MM月}与{Date.AddYears(-1):yyyy年MM月}");
                sheet.GetCell(2, 2).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 2).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 6).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 6).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 10).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 10).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 3).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 3).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[MONTH]", Date.Month.ToString());
            }
            else
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date.Year}年1-{Date.Month}月与{Date.AddYears(-1).Year}年1-{Date.Month}月");
                sheet.GetCell(2, 2).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 2).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 6).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 6).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 10).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 10).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 3).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 3).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 7).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 11).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[MONTH]", $"1-{Date.Month}");
            }
            // 行高
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).RowHeight = 27;
            // 边框
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).StyleLine();
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每月2医技科室收入数据同期比表
        /// 导出：医技科室收入数据同期比表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MYYZCXBB2(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '判断
            var BMonth = bool.Parse(com.GetParam(9));
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            if (BMonth)
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date:yyyy年MM月}与{Date.AddYears(-1):yyyy年MM月}");
                sheet.GetCell(2, 3).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 3).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 4).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 4).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 8).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 8).SetCellReplace("[MONTH]", Date.Month.ToString());
                sheet.GetCell(2, 12).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 12).SetCellReplace("[MONTH]", Date.Month.ToString());
            }
            else
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date.Year}年1-{Date.Month}月与{Date.AddYears(-1).Year}年1-{Date.Month}月");
                sheet.GetCell(2, 3).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 3).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 7).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 7).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 11).SetCellReplace("[YEAR1]", Date.AddYears(-1).Year.ToString());
                sheet.GetCell(2, 11).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 4).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 4).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 8).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 8).SetCellReplace("[MONTH]", $"1-{Date.Month}");
                sheet.GetCell(2, 12).SetCellReplace("[YEAR2]", Date.Year.ToString());
                sheet.GetCell(2, 12).SetCellReplace("[MONTH]", $"1-{Date.Month}");
            }
            // 行高
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).RowHeight = 32;
            // 边框
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).StyleLine();
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每月4每月手术人数表
        /// 导出：每月手术人数表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MYYZCXBB4(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '判断
            var BMonth = bool.Parse(com.GetParam(9));
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            if (BMonth)
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date.Year}年{Date.Month}月");
            }
            else
            {
                sheet.GetCellFirst().SetCellReplace("[DATE]", $"{Date.Year}年1-{Date.Month}月");
            }
            sheet.GetCell(2, 1).SetCellReplace("[NUM]", Data.Rows[Data.Rows.Count - 1][5].ToString());
            // 行高
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).RowHeight = 32;
            // 边框
            sheet.GetCell(RowBeginIndex, 1, RowEndIndex, Data.Columns.Count).StyleLine();
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex);
            // 合计红
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                if (Data.Rows[i][1].ToString().IsNullOrWhiteSpace())
                {
                    //字体红色加粗
                    sheet.GetRow(RowBeginIndex + i).StyleFontColorRed().StyleFontIsBold(true);
                }
            }
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        /// <summary>
        /// 模板：每月3门急诊数据同期比表
        /// 导出：门急诊数据同期比表.xlsx
        /// 参数
        /// 0：Exe地址
        /// 1：Bin地址
        /// 2：模板地址
        /// 3：保存地址
        /// 4：模板文件名
        /// 5：保存文件名
        /// 6：查询时间
        /// 7：数据导入开始行
        /// </summary>
        /// <param name="com"></param>
        public static void MYYZCXBB3(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            var PathSave = com.GetParam(3);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '保存图片地址
            var PathImageSave = com.GetParam(8);
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            var beginrow = 3;
            //设置单元格数据
            sheet.GetCell(beginrow, 1).SetCell(Date.AddYears(-1).ToString("yyyy年MM月"));
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][1].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][4].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][7].ToString());
            beginrow++;
            sheet.GetCell(beginrow, 1).SetCell(Date.ToString("yyyy年MM月"));
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][2].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][5].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][8].ToString());
            beginrow++;
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][3].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][6].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][9].ToString());
            beginrow += 2;
            int iii = 9;
            sheet.GetCell(beginrow, 1).SetCell($"{Date.AddYears(-1).Year}年1-{Date.Month}月");
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][1 + iii].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][4 + iii].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][7 + iii].ToString());
            beginrow++;
            sheet.GetCell(beginrow, 1).SetCell($"{Date.Year}年1-{Date.Month}月");
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][2 + iii].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][5 + iii].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][8 + iii].ToString());
            beginrow++;
            sheet.GetCell(beginrow, 2).SetCellNumber(Data.Rows[0][3 + iii].ToString());
            sheet.GetCell(beginrow, 3).SetCellNumber(Data.Rows[0][6 + iii].ToString());
            sheet.GetCell(beginrow, 4).SetCellNumber(Data.Rows[0][9 + iii].ToString());
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 保存图片
            var PathSaveImage = Path.ChangeExtension(Path.Combine(PathImageSave, NameSave), "png");
            Helper.SaveBmp(PathSaveImage, sheet);
        }
        #endregion

        #region 每月洛轴医保卡帐户
        public static void MYLZYBKZH(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
            // 无数据
            if (Data.Rows.Count <= 0)
                return;
            // 'Exe地址
            var PathExe = com.GetParam(0);
            // 'Bin地址
            var PathBin = com.GetParam(1);
            // '模板地址
            var PathTemplate = com.GetParam(2);
            // '保存地址
            //var PathSave = com.GetParam(3);
            var PathSave = com.GetParam(8);
            // '模板文件名
            var NameTemplate = com.GetParam(4);
            // '保存文件名
            var NameSave = com.GetParam(5);
            // '查询时间
            var Date = DateTime.ParseExact(com.GetParam(6), "yyyyMMdd", CultureInfo.CurrentCulture);
            // '数据导入开始行
            var RowBeginIndex = int.Parse(com.GetParam(7));
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            // 拷贝备份
            File.Copy(Path.Combine(PathTemplate, NameTemplate), Path.ChangeExtension(Path.Combine(PathTemplate, NameTemplate), "1.xlsx"), true);
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 计算位置
            var ColumnReginIndex = 3 + Date.Month * 2;
            // 设置单元格表头
            sheet.GetCell(1, ColumnReginIndex - 1).SetCell($"{Date:yyyy年MM月}帐户支出");
            sheet.GetCell(1, ColumnReginIndex).SetCell($"{Date:yyyy年MM月}帐户余额");
            // 导出数据到Excel
            sheet.InsertDataTable(Data, false, RowBeginIndex, ColumnReginIndex);
            // 计算列支出
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                sheet.GetCell(i + RowBeginIndex, ColumnReginIndex - 1).Text = $"={sheet.GetCell(i + RowBeginIndex, ColumnReginIndex - 2).RangeAddress}-{sheet.GetCell(i + RowBeginIndex, ColumnReginIndex).RangeAddress}";
            }
            // 计算合计
            sheet.GetCell(RowEndIndex + 1, ColumnReginIndex).Text = $"=SUM({sheet.GetCell(RowBeginIndex, ColumnReginIndex).RangeAddress}:{sheet.GetCell(RowEndIndex, ColumnReginIndex).RangeAddress})";
            sheet.GetCell(RowEndIndex + 1, ColumnReginIndex - 1).Text = $"={sheet.GetCell(RowEndIndex + 1, ColumnReginIndex - 2).RangeAddress}-{sheet.GetCell(RowEndIndex + 1, ColumnReginIndex).RangeAddress}";
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            // 替换模板
            File.Copy(Path.Combine(PathSave, NameSave), Path.Combine(PathTemplate, NameTemplate), true);
        }
        #endregion

    }
}
