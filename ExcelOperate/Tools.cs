using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text.Json;

namespace WedoExpress
{

    class Utils
    {
        /// <summary>
        /// 修复客户编号
        /// </summary>
        /// <param name="x"></param>
        /// <returns></returns>
        static public string remakeNumber(string x)
        {
            x = x.Replace(" ", "").ToLower();
            Match match = Regex.Match(x, @"^cn0{2,6}(\d+)$");

            if (match.Success)
            {
                if (match.Length != 10)
                {
                    string id = match.Groups[1].Value;
                    string fillZero = new string('0', 8 - id.Length);
                    return string.Format("cn{0}{1}", fillZero, id);
                }
                else
                {
                    return x;
                }
            }
            else
            {
                match = Regex.Match(x, @"^\d+$");
                if (match.Success)
                {
                    string id = match.Value;
                    string fillZero = new string('0', 8 - id.Length);
                    return string.Format("cn{0}{1}", fillZero, id);
                }
                else
                {
                    return "";
                }
            }


        }

        /// <summary>
        /// 获取文件夹路径
        /// </summary>
        /// <param name="tip">提示语句</param>
        /// <returns></returns>
        static public string GetSourceFolder(string tip)
        {
            Console.Write(tip);
            string input = Console.ReadLine();
            if (input == "") { return ""; }
            string source = Path.GetFullPath(input);
            if (!Directory.Exists(source))
            {
                Console.WriteLine("输入的文件夹不存在,请重新输入...");
                return GetSourceFolder(tip);
            }
            return source;
        }

        /// <summary>
        /// 获取文件路径
        /// </summary>
        /// <param name="tip">提示语句</param>
        /// <returns></returns>
        static public string GetSourceFile(string tip)
        {
            Console.Write(tip);
            string input = Console.ReadLine();
            if (input == "") { return ""; }
            string source = Path.GetFullPath(input);
            if (!File.Exists(source))
            {
                Console.WriteLine("输入的文件不存在,请重新输入...");
                return GetSourceFile(tip);
            }

            return source;
        }

        /// <summary>
        /// 获取文件夹下第一个包含标记的文件名
        /// </summary>
        /// <param name="folder">文件夹路径</param>
        /// <param name="mark">目标文件名包含的字符串</param>
        /// <returns></returns>
        static public string GetFileFullName(string folder, string mark)
        {
            DirectoryInfo root = new DirectoryInfo(folder);
            FileInfo[] files = root.GetFiles();
            foreach (var item in files)
            {
                if (item.Name.Contains(mark))
                {
                    return item.FullName;
                }
            }
            return null;
        }

        /// <summary>
        /// 将整数转换为Excel的列名
        /// </summary>
        /// <param name="index">从0开始计算的列索引</param>
        /// <returns></returns>
        static public string GetColumnLetter(int index)
        {
            int div = index / 26, mod = index % 26;
            string letter = ((char)(mod + (int)'A')).ToString();
            while (div > 0)
            {
                index -= 26;
                div = index / 26;
                mod = index % 26;
                letter += ((char)(mod + (int)'A')).ToString();
            }
            return letter;
        }

        static public string DicToCookieHeader(Dictionary<string, string> dic)
        {
            List<string> cookies = new List<string>();
            foreach (var item in dic)
            {
                cookies.Add(item.Key + "=" + item.Value);
            }
            return string.Join("; ", cookies);
        }

        static public void CookieHeaderToDic(string cookie)
        {
            foreach (var item in cookie.Split("; "))
            {
                string[] a = item.Split("=", 2);
                if ((a[0] == "com.ups.ims.lasso.sDataLassoFeb19") || (a[0] == "ups_language_preference") || (a[0] == "bm_sz") || (a[0] == "utag_main"))
                    Console.WriteLine("{" + string.Format("\"{0}\",\"{1}\"", a[0], a[1]) + "},");
            }
        }
    }

}