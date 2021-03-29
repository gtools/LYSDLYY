using System;
using System.IO;
using GTSharp;
using GTSharp.GTApp;
using Spire.Xls;

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
            try
            {
                // bin文件的地址
                var pathbin = args.Length == 0 ? "" : args[0];
                //var pathbin = @"C:\resource\金网运营管理系统\GTSharp\Bin\67e33f08-3b94-4afb-83e5-063c559dfa76.bin";
                //var pathbin = @"D:\VS\LYSDLYYWX\LYSDLYYWX\bin\Debug\GTSharp\Bin\1a28113e-67d4-4e6a-9824-96b02bb23705.bin";
                //fa6792dc-faa0-4295-b06f-d7bab6392316.bin

                // 转化为对象
                //GTDataFile datafile = GTSharp.Core.SerializeHelper.FileTObje<GTDataFile>(pathbin);

                //// 执行数据
                //GTHelper.Save(datafile);
                //GTSharp.Core.Log.Log4Net.Info("成功！");
                ////关闭
                //Close();
                //return;


                //pathbin = @"D:\VS\LYSDLYYWX\LYSDLYYWX\bin\Debug\GTSharp\Bin\224c71d9-61be-4c7a-ac43-2106af293fcb.bin";
                //pathbin = @"C:\Users\Administrator\Desktop\0f8783e4-6966-4f4e-8bb5-03092ea5c308.bin";
                // 转化为对象
                ClassCOM com = GTSharp.Core.SerializeHelper.FileTObje<ClassCOM>(pathbin);


                //switch (datafile.Command)
                switch (com.ComName)
                {
                    case "每日1科室在院人数一览表":
                        AnalysisReport.MRYYCXBB1(com);
                        //AnalysisReport.MR1(GTHelper.Save(datafile));
                        break;
                    case "每日2按手术时间统计手术人数表":
                        AnalysisReport.MRYYCXBB2(com);
                        break;
                    case "每日3在院危重病人患者明细表":
                        AnalysisReport.MRYYCXBB3(com);
                        break;
                    case "每日4在院I级护理患者明细表":
                        AnalysisReport.MRYYCXBB3(com);
                        break;
                    case "每日6在院护理无患者明细表":
                        AnalysisReport.MRYYCXBB3(com);
                        break;
                    case "每日5主要业务数据表":
                        AnalysisReport.MRYYCXBB5(com);
                        break;
                    case "每日8主要业务数据表":
                        AnalysisReport.MRYYCXBB8(com);
                        break;
                    case "每日9科室在院人数一览表":
                        AnalysisReport.MRYYCXBB9(com);
                        break;
                    case "每日10门诊退费明细":
                        AnalysisReport.MRYYCXBB10(com);
                        break;
                    case "每日11门诊日志登记表":
                        AnalysisReport.MRYYCXBB11(com);
                        break;
                    case "每日12门诊日志汇总表":
                        AnalysisReport.MRYYCXBB12(com);
                        break;
                    case "每日13门诊疑似胸痛患者列表":
                        AnalysisReport.MRYYCXBB13(com);
                        break;
                    case "全院未交病历":
                        QYWTJBL.WeekReport1(com);
                        break;
                    case "DR报告打印":
                        FSK.DRBGDY(com);
                        break;
                    case "每周院长查询报表":
                        AnalysisReport.MZYZCXBB1(com);
                        break;
                    case "每周2主要业务数据表":
                        AnalysisReport.MZYZCXBB2(com);
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
                    case "每月洛轴医保卡帐户":
                        AnalysisReport.MYLZYBKZH(com);
                        break;
                    case "每月1住院主要业务数据同期比表":
                        AnalysisReport.MYYZCXBB1(com);
                        break;
                    case "每月2医技科室收入数据同期比表":
                        AnalysisReport.MYYZCXBB2(com);
                        break;
                    case "每月4每月手术人数表":
                        AnalysisReport.MYYZCXBB4(com);
                        break;
                    case "每月3门急诊数据同期比表":
                        AnalysisReport.MYYZCXBB3(com);
                        break;
                    case "每月5主要业务数据表":
                        AnalysisReport.MYYZCXBB5(com);
                        break;
                    case "每月6主要业务数据表":
                        AnalysisReport.MYYZCXBB6(com);
                        break;
                    case "每月在院人数":
                        Class1.myzyrs(com);
                        break;
                    case "每日核酸检测信息":
                        Class1.MRHSJCXX(com);
                        break;
                    case "每日核酸检测信息1":
                        Class1.MRHSJCXX1(com);
                        break;
                    case "每日核酸检测信息2":
                        Class1.MRHSJCXX2(com);
                        break;
                    case "每日核酸检测信息3":
                        Class1.MRHSJCXX3(com);
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