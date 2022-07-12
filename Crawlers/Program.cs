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
        /// 生成UPS签收单
        /// </summary>
        static public void MakeJackUpsPDF()
        {
            //声明变量
            string dataPath, saveDir, username, password;

            //获取路径
            dataPath = Utils.GetSourceFile("请输入单号文档：");
            if (dataPath == "") { return; }//如果为空字符串，代表想要退出功能
            saveDir = Path.Join(Directory.GetParent(dataPath).FullName, "UPS签收单导出", DateTime.Today.ToLongDateString());
            if (!Directory.Exists(saveDir)) { Directory.CreateDirectory(saveDir); }
            string[] lines = File.ReadAllLines(dataPath);
            username = lines[0].Replace("userId:", "").Trim();
            password = lines[1].Replace("password:", "").Trim();

            Console.WriteLine("\n正在登录...\n");
            UPSCrawler client = new UPSCrawler(saveDir);
            client.Login(username, password);



            Console.WriteLine("\n正在输出签收单，请不要关闭此窗口...\n");

            //创建计时器计算运行时间
            Stopwatch watch = new Stopwatch();
            watch.Start();

            //运行的主体函数
            var list = new List<Task>();
            for (int i = 2; i < lines.Length; i++)
            {
                if (!lines[i].Equals(""))
                    list.Add(client.MakeUpsPDF(lines[i], saveDir));
            }
            Task.WaitAll(list.ToArray());

            //停止计时器
            watch.Stop();

            //运行报告
            Console.WriteLine("\n签收单输出完毕，耗时 {0} ，按任意键返回主页...", watch.Elapsed);
            Console.ReadKey();

        }

        /// <summary>
        /// 阿里国际联系方式爬虫
        /// </summary>
        static public void AliApplierCrawler()
        {
            if (!File.Exists("login_cookies.txt"))
            {
                Console.WriteLine("login_cookies.txt 文件不存在，按任意键返回主页。");
                Console.ReadKey();
                return;
            }

            AliCrawler ali = new AliCrawler(File.ReadAllText("login_cookies.txt"));

            Console.Write("请输入供应商关键词:");
            string keyword = Console.ReadLine();//"electric bicycle"

            string searchUrl = ali.SearchKeyword(keyword);

            // 保存本次搜索得到的所有供应商联系方式的网址
            List<string> contactList = new List<string>();
            using (var file = new StreamWriter(keyword + "_list.txt"))
            {
                List<string> urls = new List<string>();
                var searchRes = ali.ParseSearchPage(searchUrl);
                do
                {
                    foreach (var contactUrl in searchRes.contacts)
                    {
                        urls.Add(contactUrl);
                    }
                    searchRes = ali.ParseSearchPage(searchRes.next);
                } while (searchRes.next != null);
                HashSet<string> searchSet = new HashSet<string>(urls);
                Console.WriteLine("共搜索出{0}，去重排除掉:{1}个", urls.Count, urls.Count - searchSet.Count);
                foreach (var item in searchSet)
                {
                    file.WriteLine(item);
                    contactList.Add(item);
                }
            }


            var values = new List<Dictionary<string, string>>();
            int repeatcount = 0;
            for (int i = 0; i < contactList.Count; i++)
            {
                Console.SetCursorPosition(0, Console.CursorTop);
                Console.Write("当前处理进度{0}/{1}", i + 1, contactList.Count);

                string contactUrl = contactList[i];

                try
                {
                    var ContactRes = ali.ReadContactPage(contactUrl);
                    if (ContactRes == (null, null))
                    {
                        Console.WriteLine("\nmiss:" + contactUrl);
                        continue;
                    }

                    values.Add(ali.ParseContactPage(contactUrl, ContactRes.contactTable, ContactRes.encryptAccountId));
                    // System.Threading.Thread.Sleep(500);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    repeatcount++;
                    if (repeatcount < 2)
                    {
                        i--;
                        System.Threading.Thread.Sleep(5000);
                        Console.WriteLine("\n再次尝试:" + contactUrl);
                    }
                }

            }
            MiniExcel.SaveAs(keyword + "供应商联系方式.xlsx", values);

            return;
        }

    }


    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("项目已上传至Github:https://github.com/yuyuko-C/WedoTools");
                Console.WriteLine("可选功能:1.UPS签收单导出   2.阿里供应商导出");
                Console.Write("请输入你要使用的功能编号:");
                string input = Console.ReadLine();
                Console.Clear();//获取功能后清空屏幕
                switch (input)
                {
                    case "1":
                        Console.WriteLine("@欢迎使用UPS签收单导出功能\n");
                        WedoTool.MakeJackUpsPDF();
                        break;
                    case "2":
                        Console.WriteLine("@欢迎使用阿里供应商导出功能\n");
                        WedoTool.AliApplierCrawler();
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
