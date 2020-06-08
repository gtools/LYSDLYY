using GTSharp;
using GTSharp.Extension;
using Spire.Xls;
using System;
using System.Globalization;
using System.IO;

namespace LYSDLYY
{
    class Class1
    {
        /// <summary>
        /// 河南省医疗服务恢复情况监测周报表
        /// </summary>
        public static void hnsylfwhfqkjczbb(ClassCOM com)
        {
            // 数据
            var Data = com.Data.Tables[0].Copy();
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
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            //// 设置单元格日期
            //sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月"));
            //// 设置单元格计数
            //sheet.GetCellFirst().SetCellReplace("[NUM]", (Helper.GetWeekNumInMonth(Date) - 1).ToString());
            //sheet.DataTableToExcel(Data, RowBeginIndex);

            //设置单元格数据
            //日期  2020年 03 月 23 日——   03月  29  日
            var temp = Date.AddDays(-6).ToString("yyyy年 MM 月 dd 日") + "——   " + Date.ToString("MM月  dd  日");
            sheet.GetCell(3, 1).SetCellReplace("[DATE]", temp);
            //门诊人次
            sheet.SetCellValue(6, 3, Data.Rows[0][0].ToString());
            sheet.SetCellValue(6, 4, Data.Rows[1][0].ToString());
            //急诊人次
            sheet.SetCellValue(6, 6, Data.Rows[0][1].ToString());
            sheet.SetCellValue(6, 7, Data.Rows[1][1].ToString());
            //住院人次
            sheet.SetCellValue(6, 9, Data.Rows[0][2].ToString());
            sheet.SetCellValue(6, 10, Data.Rows[1][2].ToString());
            //出院人次
            sheet.SetCellValue(6, 12, Data.Rows[0][3].ToString());
            sheet.SetCellValue(6, 13, Data.Rows[1][3].ToString());
            //手术台次
            sheet.SetCellValue(6, 15, Data.Rows[0][4].ToString());
            sheet.SetCellValue(6, 16, Data.Rows[1][4].ToString());
            book.SaveToFile(Path.Combine(PathSave, NameSave));
            sheet.Dispose();
            book.Dispose();
        }

        /// <summary>
        /// 入院人数和门急诊就诊人数
        /// </summary>
        /// <param name="com"></param>
        public static void ryrshmjzjzrs(ClassCOM com)
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
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月dd日"));
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex, true);
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
        }

        /// <summary>
        /// 心血管疾病病人信息
        /// </summary>
        /// <param name="com"></param>
        public static void xxgjbbrxx(ClassCOM com)
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
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月"));
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex, true);
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
        }

        /// <summary>
        /// 每月在院人数
        /// </summary>
        /// <param name="com"></param>
        public static void myzyrs(ClassCOM com)
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
            var book = new Workbook();
            book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
            var sheet = book.Worksheets[0];
            // 设置单元格日期
            sheet.GetCellFirst().SetCellReplace("[DATE]", Date.ToString("yyyy年MM月"));
            // 导出数据到Excel
            sheet.DataTableToExcel(Data, RowBeginIndex, true);
            // 保存
            book.SaveToFile(Path.Combine(PathSave, NameSave));
        }
    }
}
