using System;
using System.IO;
using System.Threading;
using GTSharp;

namespace LYSDLYY
{
    public class Program
    {
        /// <summary>
        /// 主函数
        /// </summary>
        /// <param name="args">需要传入的数据</param>
        public static void Main(string[] args)
        {
            Thread.Sleep(500);
            try
            {
                // bin文件的地址
                var pathbin = @"C:\resource\金网运营管理系统\GTSharp\Bin\d9ca6fc0-0485-45ff-a599-f67cf7f91dd7.bin";
                //var pathbin = args.Length == 0 ? "" : args[0];
                // 转化为对象
                ClassCOM com = GTSharp.Core.SerializeHelper.FileTObje<ClassCOM>(pathbin);
                switch (com.ComName)
                {
                    case "科室在院人数一览表":
                        AnalysisReport.MRYYCXBB1(com);
                        break;
                    case "按手术时间统计手术人数表":
                        AnalysisReport.MRYYCXBB2(com);
                        break;
                    case "在院危重病人患者明细表":
                        AnalysisReport.MRYYCXBB3(com);
                        break;
                    /*
                case "在院I级护理患者明细表":
                    AnalysisReport.MRYYCXBB4(com);
                    break;
                    */
                    case "主要业务数据表":
                        AnalysisReport.MRYYCXBB5(com);
                        break;
                    case "全院未交病历":
                        QYWTJBL.WeekReport1(com);
                        break;
                    case "DR报告打印":
                        FSK.DRBGDY(com);
                        break;
                    case "每周院长查询报表":
                        AnalysisReport.MZYZCXBB(com);
                        break;
                    case "删除多余数据":
                        DeleteBin(Path.GetDirectoryName(pathbin), 50);
                        break;
                    case "河南省医疗服务恢复情况监测周报表":
                        Class1.hnsylfwhfqkjczbb(com);
                        break;
                    case "入院人数和门急诊就诊人数":
                        Class1.ryrshmjzjzrs(com);
                        break;
                    case "心血管疾病病人信息":
                        Class1.xxgjbbrxx(com);
                        break;
                    //
                    default:
                        // 删除多余数据
                        DeleteBin(Path.GetDirectoryName(pathbin), 50);
                        GTSharp.Core.Log.Log4Net.Info("未知命令?");
                        Console.ReadLine();
                        break;
                }
                GTSharp.Core.Log.Log4Net.Info("成功！");
                //关闭
                Close();
            }
            catch (Exception ex)
            {
                GTSharp.Core.Log.Log4Net.Error(ex);
                Console.ReadLine();
            }
        }
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="count">文件数量</param>
        static void DeleteBin(string path, int count)
        {
            if (Directory.GetFiles(path).Length > count)
                Directory.Delete(path, true);
        }
        /// <summary>
        /// 关闭程序
        /// </summary>
        static void Close()
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }
    }
}