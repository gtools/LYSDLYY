using GTSharp.GTApp;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LYSDLYY
{
    /// <summary>
    /// 根据GTSharp.GTApp.GTDataFile生成数据
    /// </summary>
    public class GTHelper
    {
        /// <summary>
        /// 保存数据
        /// </summary>
        /// <param name="file">GTSharp.GTApp.GTDataFile</param>
        public static void Save(GTDataFile file)
        {
            // 文件循环
            foreach (GTFile fileitem in file.Files)
            {
                // 初始化workbook实例
                Workbook workbook = new Workbook();
                // 加载Excel文档
                workbook.LoadFromStream(new MemoryStream(fileitem.File), ExcelVersion.Version2007);
                // 获取第一个工作表
                Worksheet sheet = workbook.Worksheets[0];
                // 数据循环
                foreach (GTData dataitem in fileitem.Datas)
                {
                    // 是否导入数据
                    if (dataitem.Points.Count>0)
                    {
                        // 循环数据单元格输出
                        foreach (GTPoint item in dataitem.Points)
                        {
                            sheet.SetCellValue(item.X2, item.Y2, dataitem.Table.Rows[item.X1][item.Y1].ToString());
                        }
                    }
                    else
                    {

                        // 导出数据到Excel
                        sheet.DataTableToExcel(dataitem.Table, dataitem.RowIndexBegin);
                    }
                }
                workbook.SaveToFile(fileitem.PathSave);
            }
        }
    }
}
