﻿using GTSharp;
using Spire.Xls;
using System;
using System.Data;
using System.Linq;
using System.IO;
using System.Globalization;

namespace LYSDLYY
{
    /// <summary>
    /// 放射科
    /// </summary>
    public class FSK
    {
        public static void DRBGDY(ClassCOM com)
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
            var PathSave = com.GetParam(3);
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
            var dr = Data.AsEnumerable();
            foreach (var item in dr)
            {
                var book = new Workbook();
                book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
                var sheet = book.Worksheets[0];
                sheet.Range[5, 1].Value = item[0].ToString();
                sheet.Range[5, 2].Value = item[1].ToString();
                sheet.Range[5, 3].Value = item[2].ToString();
                sheet.Range[5, 4].Value = item[3].ToString();
                sheet.Range[5, 5].Value = item[4].ToString();
                book.SaveToFile(Path.Combine(PathSave, item[1].ToString() + ".xlsx"));
                book.PrintDocument.Print();
            }
        }
    }
}
