using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using MiniExcelLibs;

namespace WedoExpress
{

    public static class WedoTool
    {
        /// <summary>
        /// 合并表格
        /// </summary>
        static public void CombineTimeLimit()
        {
            string source = Utils.GetSourceFolder("请输入表格所在的文件夹：");
            if (source == "") { return; }//如果为空字符串，代表想要退出功能

            //创建计时器计算运行时间
            Stopwatch watch = new Stopwatch();
            watch.Start();

            //运行的主体函数
            TimeLimitExtract.Extract(source);

            watch.Stop();

            //运行报告
            Console.WriteLine("\n表格输出完毕，耗时 {0} ，按任意键返回主页...", watch.Elapsed);
            Console.ReadKey();
        }

        /// <summary>
        /// 拆分表格
        /// </summary>
        static public void ApartTimeLimit()
        {
            //声明变量
            string dataPath, matchPath;

            //获取路径
            dataPath = Utils.GetSourceFile("请输入需要拆分的表格：");
            if (dataPath == "") { return; }//如果为空字符串，代表想要退出功能
            matchPath = Utils.GetSourceFile("客户客服对照的表格：");
            if (matchPath == "") { return; }

            Console.WriteLine("\n正在输出表格，请不要关闭此窗口...\n");

            //创建计时器计算运行时间
            Stopwatch watch = new Stopwatch();
            watch.Start();

            //运行的主体函数
            TimeLimitGroupBy.GroupBy(dataPath, matchPath);

            watch.Stop();

            //运行报告
            Console.WriteLine("\n表格输出完毕，耗时 {0} ，按任意键返回主页...", watch.Elapsed);
            Console.ReadKey();
        }

    }


    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("项目已上传至Github:https://github.com/yuyuko-C/WedoTools");
                Console.WriteLine("可选功能:1.表格合并   2.表格拆分");
                Console.Write("请输入你要使用的功能编号:");
                string input = Console.ReadLine();
                Console.Clear();//获取功能后清空屏幕
                switch (input)
                {
                    case "1":
                        Console.WriteLine("@欢迎使用表格合并功能\n");
                        WedoTool.CombineTimeLimit();
                        break;
                    case "2":
                        Console.WriteLine("@欢迎使用表格拆分功能\n");
                        WedoTool.ApartTimeLimit();
                        break;
                    default:
                        Console.WriteLine("想要更多功能？一杯奶茶~");
                        Console.WriteLine("按任意键返回主页...");
                        Console.ReadKey();
                        break;
                }
                Console.Clear();
            }
        }
    }
}
