using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;
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

    class UPSPDFMaker
    {
        XFont fontHeadBottom, fontTitle, fontText, fontProject;
        PdfDocument document;
        PdfPage currentPage;
        XGraphics graphics;
        double x = 43, y = 14;//文字左上角的坐标
        public UPSPDFMaker()
        {
            try
            {
                fontHeadBottom = new XFont("Arial", 8, XFontStyle.Regular);
            }
            catch
            {
                Console.WriteLine("aa???");
            }
            // fontHeadBottom = new XFont("Arial", 8, XFontStyle.Regular);
            fontTitle = new XFont("Tahoma", 19, XFontStyle.Regular);
            fontText = new XFont("Tahoma", 8.35, XFontStyle.Regular);
            fontProject = new XFont("Tahoma", 8.35, XFontStyle.Bold);
            document = new PdfDocument();
            currentPage = document.AddPage();
            currentPage.Size = PdfSharpCore.PageSize.A4;
            graphics = XGraphics.FromPdfPage(currentPage);
        }

        public void DrawPargraph(string text, XFont font, double yMove, double yMoveLine = 1)
        {
            string[] array = text.Split("\n");
            if (array.Length == 1)
            {
                graphics.DrawString(text, font, XBrushes.Black, x, y, XStringFormats.TopLeft);
            }
            else if (array.Length > 1)
            {
                for (int i = 0; i < array.Length; i++)
                {
                    if (i < array.Length - 1)
                        DrawPargraph(array[i], fontText, yMoveLine);
                    else
                        DrawPargraph(array[i], fontText, -font.Size);
                }
            }
            y += font.Size + yMove;

        }

        public void DrawGroup(string project, string text, double yMove, XImage image = null)
        {
            DrawPargraph(project, fontProject, 5);

            DrawPargraph(text, fontText, 2);

            if (image != null)
            {
                graphics.DrawImage(image, x, y, 120, 40);
                y += 40;
            }
            y += yMove;
        }

        public PdfDocument MakeUp(Dictionary<string, object> info)
        {
            //绘制页眉页脚
            graphics.DrawString(DateTime.Today.ToShortDateString(), fontHeadBottom, XBrushes.Black, 25, y, XStringFormats.TopLeft);
            graphics.DrawString("Tracking | UPS - Canada", fontHeadBottom, XBrushes.Black, currentPage.Width / 2, y, XStringFormats.TopCenter);
            graphics.DrawString("1/1", fontHeadBottom, XBrushes.Black, currentPage.Width - 38, currentPage.Height - 26, XStringFormats.TopLeft);
            y += fontHeadBottom.Size + 16;

            //绘制标题
            DrawPargraph("Proof of Delivery", fontTitle, 25);

            //绘制问候
            DrawPargraph("Dear Customer,", fontText, 7);
            DrawPargraph("This notice serves as proof of delivery for the shipment listed below.", fontText, 12);

            //绘制信息
            DrawGroup("Tracking Number", info["trackingNumber"].ToString(), 12);
            DrawGroup("Weight", info["weight"].ToString() + " " + info["weightUnit"].ToString(), 12);
            var service = JsonSerializer.Deserialize<Dictionary<string, object>>(info["service"].ToString());
            DrawGroup("Service", service["serviceName"].ToString().Replace("&#174;", "®"), 12);
            DrawGroup("Shipped / Billed On", info["shippedOrBilledDate"].ToString(), 12);
            DrawGroup("Additional Information", "Signature Required", 12);
            DrawGroup("Delivered On", info["deliveredDate"].ToString() + " " + info["deliveredTime"].ToString(), 12);
            DrawGroup("Delivered To", ToAddressParse(info), 12);
            // DrawGroup("Received By", info["receivedBy"].ToString(), 11, XImage.FromFile(info["img"].ToString()));
            DrawGroup("Received By", info["receivedBy"].ToString(), 11);
            if (info["leftAt"].ToString() != "")
                DrawGroup("Left At", info["leftAt"].ToString(), 12);
            if (info["referenceNumbers"] != null)
            {
                List<object> list = info["referenceNumbers"] as List<object>;
                DrawGroup("Reference Number(s):", list[0].ToString(), 12);
            }

            // 绘制底部
            string end = "Thank you for giving us this opportunity to serve you. Details are only available for shipments delivered within the last 120 days. Please\n" +
            "print for your records if you require this information after 120 days.";
            DrawPargraph(end, fontText, 12, 2);

            end = "Sincerely,\n" + "UPS\n" + string.Format("Tracking results provided by UPS:{0}", info["trackedDateTime"].ToString());
            DrawPargraph(end, fontText, 8, 8);


            return document;
        }

        private string ToAddressParse(Dictionary<string, object> info)
        {
            var shipToAddress = JsonSerializer.Deserialize<Dictionary<string, object>>(info["shipToAddress"].ToString());
            List<string> list = new List<string>();

            if ((shipToAddress["companyName"] != null) && (shipToAddress["companyName"].ToString() != "")) list.Add(shipToAddress["companyName"].ToString());
            if ((shipToAddress["attentionName"] != null) && (shipToAddress["attentionName"].ToString() != "")) list.Add(shipToAddress["attentionName"].ToString());
            string name = string.Join<string>("\n", list); list.Clear();

            if ((shipToAddress["streetAddress1"] != null) && (shipToAddress["streetAddress1"].ToString() != "")) list.Add(shipToAddress["streetAddress1"].ToString());
            if ((shipToAddress["streetAddress2"] != null) && (shipToAddress["streetAddress2"].ToString() != "")) list.Add(shipToAddress["streetAddress2"].ToString());
            if ((shipToAddress["streetAddress3"] != null) && (shipToAddress["streetAddress3"].ToString() != "")) list.Add(shipToAddress["streetAddress3"].ToString());
            string address = string.Join<string>("\n", list); list.Clear();

            if ((shipToAddress["city"] != null) && (shipToAddress["city"].ToString() != "")) list.Add(shipToAddress["city"].ToString());
            if ((shipToAddress["state"] != null) && (shipToAddress["state"].ToString() != "")) list.Add(shipToAddress["state"].ToString());
            if ((shipToAddress["province"] != null) && (shipToAddress["province"].ToString() != "")) list.Add(shipToAddress["province"].ToString());
            if ((shipToAddress["zipCode"] != null) && (shipToAddress["zipCode"].ToString() != "")) list.Add(shipToAddress["zipCode"].ToString());
            if ((shipToAddress["country"] != null) && (shipToAddress["country"].ToString() != "")) list.Add(shipToAddress["country"].ToString());
            string others = string.Join<string>("\n", list); list.Clear();

            if (name != "") list.Add(name);
            if (address != "") list.Add(address);
            if (others != "") list.Add(others);

            return string.Join<string>("\n", list);
        }

    }

}