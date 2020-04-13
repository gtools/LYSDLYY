using GTSharp.IO;
using Spire.Xls;
using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace LYSDLYY
{
    /// <summary>
    /// 帮助类
    /// </summary>
    public static class Helper
    {
        /// <summary>
        /// 获取最后一行的行号
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="rowIndexBegin"></param>
        /// <returns></returns>
        public static int GetRowIndexEnd(this DataTable dt, int rowIndexBegin)
        {
            return dt.Rows.Count + rowIndexBegin - 1;
        }
        /// <summary>
        /// 查询替换文字
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="oldstr">查找的文本</param>
        /// <param name="newstr">替换的文本</param>
        public static void FindAllString(Worksheet sheet, string oldstr, string newstr)
        {
            //查找字符串“紧张”
            CellRange[] ranges = sheet.FindAllString(oldstr, false, false);
            foreach (CellRange range in ranges)
            {
                //使用 “充足”替换
                range.Text = range.Text.Replace(oldstr, newstr);
                //设置高亮显示颜色
                //range.Style.Color = Color.Yellow;
            }
        }

        /// <summary>
        /// 剪裁 -- 用GDI+
        /// </summary>
        /// <param name="b">原始Bitmap</param>
        /// <param name="StartX">开始坐标X</param>
        /// <param name="StartY">开始坐标Y</param>
        /// <param name="iWidth">宽度</param>
        /// <param name="iHeight">高度</param>
        /// <returns>剪裁后的Bitmap</returns>
        public static Bitmap KiCut(Bitmap b, int StartX, int StartY, int iWidth, int iHeight)
        {
            if (b == null)
            {
                return null;
            }
            int w = b.Width;
            int h = b.Height;
            if (StartX >= w || StartY >= h)
            {
                return null;
            }
            if (StartX + iWidth > w)
            {
                iWidth = w - StartX;
            }
            if (StartY + iHeight > h)
            {
                iHeight = h - StartY;
            }
            try
            {
                Bitmap bmpOut = new Bitmap(iWidth, iHeight, PixelFormat.Format24bppRgb);
                Graphics g = Graphics.FromImage(bmpOut);
                g.DrawImage(b, new Rectangle(0, 0, iWidth, iHeight), new Rectangle(StartX, StartY, iWidth, iHeight), GraphicsUnit.Pixel);
                g.Dispose();
                return bmpOut;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取日期是月中第几周
        /// </summary>
        /// <param name="daytime"></param>
        /// <returns></returns>
        public static int GetWeekNumInMonth(DateTime daytime)
        {
            int dayInMonth = daytime.Day;
            //本月第一天
            DateTime firstDay = daytime.AddDays(1 - daytime.Day);
            //本月第一天是周几
            int weekday = (int)firstDay.DayOfWeek == 0 ? 7 : (int)firstDay.DayOfWeek;
            //本月第一周有几天
            int firstWeekEndDay = 7 - (weekday - 1);
            //当前日期和第一周之差
            int diffday = dayInMonth - firstWeekEndDay;
            diffday = diffday > 0 ? diffday : 1;
            //当前是第几周,如果整除7就减一天
            int WeekNumInMonth = ((diffday % 7) == 0
             ? (diffday / 7 - 1)
             : (diffday / 7)) + 1 + (dayInMonth > firstWeekEndDay ? 1 : 0);
            return WeekNumInMonth;
        }

        /// <summary>
        /// 保存白边剪裁过
        /// </summary>
        public static void SaveBmp(string PathSaveImage, Worksheet sheet)
        {
            DirectoryHelper.Create(Path.GetDirectoryName(PathSaveImage));
            sheet.SaveToImage(PathSaveImage, ImageFormat.Png);
            // 处理白边
            Bitmap bitmap = new Bitmap(PathSaveImage);
            Bitmap bitmap1 = KiCut(bitmap, 66, 66, bitmap.Width - 66 - 66, bitmap.Height - 66 - 66);
            bitmap.Dispose();
            bitmap1.Save(PathSaveImage);
        }
    }
}
