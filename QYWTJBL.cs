using GTSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;
using Spire.Xls;
using System.Globalization;

namespace LYSDLYY
{
    /// <summary>
    /// 全院未提交病历
    /// </summary>
    public class QYWTJBL
    {
        /// <summary>
        /// 导出模板：每周1全院未交病历.xls
        /// 参数
        /// 0：DataTable数据
        /// 0：执行命令
        /// 1：写入数据开始行
        /// 2：日期增加天数
        /// 3：模板文件路径
        /// 4：保存文件路径
        /// </summary>
        /// <param name="com"></param>
        public static void WeekReport1(ClassCOM com)
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
            // '数据导入结束行
            var RowEndIndex = RowBeginIndex + Data.Rows.Count - 1;
            // 科室
            List<string> depts = GetDepts();
            foreach (var item in depts)
            {
                var book = new Workbook();
                book.LoadFromFile(Path.Combine(PathTemplate, NameTemplate));
                var sheet = book.Worksheets[0];
                var dr = Data.AsEnumerable().Where(t => t["科室"].ToString() == item);
                if (dr.Count() > 0)
                {
                    DataTable dt = dr.OrderBy(t => t["住院号"].ToString()).CopyToDataTable();
                    RowEndIndex = RowBeginIndex + dt.Rows.Count - 1;
                    dt.Columns.Remove("科室");
                    sheet.InsertDataTable(dt, false, RowBeginIndex, 1);
                    var call = sheet.Range[RowBeginIndex, 1, RowEndIndex, dt.Columns.Count];
                    call.StyleLine();
                    Helper.FindAllString(sheet, "[DEPT]", item);
                    Helper.FindAllString(sheet, "[NUM]", dt.Rows.Count.ToString());
                    Helper.FindAllString(sheet, "[DATE]", DateTime.Now.ToString("yyyy-MM-dd"));
                    book.SaveToFile(Path.Combine(PathSave, item + ".xlsx"));
                    book.PrintDocument.Print();
                }
                //book.PrintDocument.PrintController = new StandardPrintController();
            }
        }
        public static List<string> GetDepts()
        {
            return new List<string>()
            {
                "呼吸内科、职业病病区",
                "老年病、普内科病区",
                "神经内科病区",
                "神经内科二病区",
                "心血管内科病区",
                "消化科、肿瘤科病区",
                "普外科、胸外科病区",
                "泌尿外科病区",
                "骨科二病区",
                "新五官科病区",
                "普外科、脑外科病区",
                "骨科病区",
                "新儿科康复病区",
                "新儿科病区",
                "儿科",
                "心血管内科二病区",
                "神经内科二病区",
                "新妇产科病区",
                "重症医学科(ICU)",
                "内六科"
            };
        }
    }
}
