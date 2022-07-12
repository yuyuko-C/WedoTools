using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Attributes;
using WedoExpress;


namespace WedoExpress
{

    public class ExcelFileReader : IDisposable
    {
        public IEnumerable<dynamic> rows;
        public ICollection<string> columns;
        private FileStream stream;
        public string fileMark, sheetName;

        public ExcelFileReader(string folder, string fileMark)
        {
            string path = Utils.GetFileFullName(folder, fileMark);
            if (path == null)
            {
                Console.WriteLine("未找到表格文件。\n 文件夹：{0}\n文件特征:{1}", folder, fileMark);
                return;
            }
            stream = File.OpenRead(path);
            this.fileMark = fileMark;
        }

        public ExcelFileReader(string path)
        {
            if (!File.Exists(path))
            {
                Console.WriteLine("未找到表格文件。{0}", path);
                return;
            }
            stream = File.OpenRead(path);
        }

        public void ReadSheet(string sheetName = null, int startRow = 0)
        {
            if (fileMark == null)
                Console.Write("开始读取 《{0}》表 —— ", sheetName);
            else
                Console.Write("开始读取 {0} 的《{1}》表 —— ", fileMark, sheetName);

            var config = new OpenXmlConfiguration()
            {
                FillMergedCells = true
            };
            string startCell = string.Format("A{0}", startRow);
            rows = stream.Query(useHeaderRow: true, sheetName: sheetName, startCell: startCell, configuration: config);
            columns = stream.GetColumns(useHeaderRow: true, sheetName: sheetName, startCell: startCell, configuration: config);

            if (fileMark == null)
                Console.WriteLine("《{0}》表读取完成", sheetName);
            else
                Console.WriteLine("{0} 的《{1}》表读取完成", fileMark, sheetName);


            this.sheetName = sheetName;
        }


        public void Dispose()
        {
            stream.Dispose();
        }
    }


    public class TimeLimitExtract
    {
        private static void USLine_Fast(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {

            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["平台/客户"],
                    ["发货日期"] = row["系统发货日期"],
                    ["渠道"] = row["渠道"],
                    ["发往国家"] = row["国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["末端上网时间"],
                    ["妥投日期"] = row["末端妥投时间"],
                    ["当前状态"] = row["末端状态"] == null ? row["头程状态"] : row["末端状态"],
                    ["备注"] = null,
                    ["头程状态"] = row["头程状态"],
                });
            }
        }

        private static void USLine_Cheap(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["客户"],
                    ["发货日期"] = row["系统发货日期"],
                    ["渠道"] = row["渠道"],
                    ["发往国家"] = row["国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["末端上网时间"],
                    ["妥投日期"] = row["末端妥投时间"],
                    ["当前状态"] = row["末端状态"] == null ? row["头程状态"] : row["末端状态"],
                    ["备注"] = null,
                    ["头程状态"] = row["头程状态"],
                });
            }
        }

        private static void USLine_Prefer(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["平台/客户"],
                    ["发货日期"] = row["系统发货日期"],
                    ["渠道"] = row["发货渠道"],
                    ["发往国家"] = row["国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["末端上网时间"],
                    ["妥投日期"] = row["末端妥投时间"],
                    ["当前状态"] = row["末端状态"],
                    ["备注"] = row["备注"],
                    ["头程状态"] = null,
                });
            }
        }

        private static void Fedex_LowWeight(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["平台/客户"],
                    ["发货日期"] = row["系统发货日期"],
                    ["渠道"] = row["发货渠道"],
                    ["发往国家"] = row["国家"],
                    ["跟踪号"] = row["最终发货单号"] != null ? row["最终发货单号"] : row["跟踪号"],
                    ["上网日期"] = row["末端上网时间"],
                    ["妥投日期"] = row["末端妥投时间"],
                    ["当前状态"] = row["末端状态"],
                    ["备注"] = null,
                    ["头程状态"] = null,
                });
            }
        }

        private static void SpecialGoods(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["平台/客户"],
                    ["发货日期"] = row["系统发货日期"],
                    ["渠道"] = row["发货渠道"],
                    ["发往国家"] = row["国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["末端上网时间"],
                    ["妥投日期"] = row["末端妥投时间"],
                    ["当前状态"] = row["末端状态"],
                    ["备注"] = row["备注"],
                    ["头程状态"] = null,
                });
            }
        }

        private static void InternalWarehouse(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["注册名"],
                    ["发货日期"] = row["系统发货时间"],
                    ["渠道"] = row["渠道"],
                    ["发往国家"] = row["发往国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["上网日期"],
                    ["妥投日期"] = row["妥投日期"],
                    ["当前状态"] = row["当前状态"],
                    ["备注"] = row["异常情况"],
                    ["头程状态"] = null,
                });
            }
        }

        private static void AboardWarehouse(ExcelFileReader reader, ref List<Dictionary<string, object>> retValues)
        {
            Console.WriteLine("开始处理 {0} 的《{1}》表", reader.fileMark, reader.sheetName);
            foreach (IDictionary<string, object> row in reader.rows)
            {
                retValues.Add(new Dictionary<string, object>()
                {
                    ["客户编号"] = row["平台账号"],
                    ["发货日期"] = row["发货时间"],
                    ["渠道"] = row["渠道"],
                    ["发往国家"] = row["发往国家"],
                    ["跟踪号"] = row["跟踪号"],
                    ["上网日期"] = row["上网时间"],
                    ["妥投日期"] = row["妥投时间"],
                    ["当前状态"] = row["跟踪状态"],
                    ["备注"] = null,
                    ["头程状态"] = null,
                });
            }
        }


        public static void Extract(string folderPath)
        {
            //对输出路径进行检查
            string outpath = Path.Combine(folderPath, "合并后的表格.xlsx");
            Console.WriteLine("完成后表格将会在：{0}\n", outpath);
            Console.WriteLine("正在检查表格路径是否被占用.........");
            while (File.Exists(outpath))
            {
                Console.WriteLine("表格路径被占用，请删除该文件后按任意键继续...");
                Console.ReadKey();
            }

            Console.WriteLine("\n路径检查完毕，正在输出表格，请不要关闭此窗口...\n");

            List<Dictionary<string, object>> retValues = new List<Dictionary<string, object>>();
            ExcelFileReader reader;

            using (reader = new ExcelFileReader(folderPath, "美国专线订单跟单数据"))
            {
                reader.ReadSheet("美国专线发货订单明细");
                USLine_Fast(reader, ref retValues);
                reader.ReadSheet("美国经济专线发货订单明细");
                USLine_Cheap(reader, ref retValues);
            }

            using (reader = new ExcelFileReader(folderPath, "美国专线小包-特惠"))
            {
                reader.ReadSheet("发货订单明细");
                USLine_Prefer(reader, ref retValues);
            }

            using (reader = new ExcelFileReader(folderPath, "联邦小货订单跟单数据"))
            {
                reader.ReadSheet("数据源");
                Fedex_LowWeight(reader, ref retValues);
            }

            using (reader = new ExcelFileReader(folderPath, "国内仓"))
            {
                reader.ReadSheet("原数据");
                InternalWarehouse(reader, ref retValues);
            }

            using (reader = new ExcelFileReader(folderPath, "海外仓"))
            {
                reader.ReadSheet("原数据");
                AboardWarehouse(reader, ref retValues);
            }

            //清除不需要的数据
            for (int i = retValues.Count - 1; i >= 0; i--)
            {
                if (retValues[i]["客户编号"] != null)
                {
                    string name = retValues[i]["客户编号"].ToString();
                    name = Utils.remakeNumber(name);
                    retValues[i]["客户编号"] = name;
                    if (name.Equals(""))
                    {
                        retValues.RemoveAt(i);
                    }
                }
                else
                {
                    retValues.RemoveAt(i);
                }

            }


            MiniExcel.SaveAs(outpath, retValues);
        }

    }


    public class TimeLimitGroupBy
    {
        private static Dictionary<string, List<IDictionary<string, object>>> GroupByColumn(IEnumerable<dynamic> retValues, string column)
        {
            Dictionary<string, List<IDictionary<string, object>>> dic = new Dictionary<string, List<IDictionary<string, object>>>();
            foreach (IDictionary<string, object> row in retValues)
            {
                string dicKey = row[column].ToString();
                if (!dic.ContainsKey(dicKey))
                {
                    dic.Add(dicKey, new List<IDictionary<string, object>>() { row });
                }
                else
                {
                    dic[dicKey].Add(row);
                }

            }
            return dic;
        }

        private static Dictionary<string, string> GetUIDMatchSeverDic(string path)
        {
            using (var file = new ExcelFileReader(path))
            {
                file.ReadSheet();
                var rows = file.rows.Cast<IDictionary<string, object>>();
                var dic = new Dictionary<string, string>();
                foreach (var row in rows)
                {
                    dic[row["客户编号"].ToString()] = row["客服专员"].ToString();
                }
                return dic;
            }
        }

        public static void GroupBy(string dataPath, string matchPath)
        {
            Dictionary<string, List<IDictionary<string, object>>> FBAgroup;
            Dictionary<string, List<IDictionary<string, object>>> PACKgroup;

            //删除已存在的时效文件夹
            string folder = Directory.GetParent(dataPath).FullName;
            foreach (var item in Directory.EnumerateDirectories(folder))
            {
                if (item.Contains("客户时效表")) { Directory.Delete(item, true); }
            }

            //客户客服字典
            Dictionary<string, string> UIDMatchSever = GetUIDMatchSeverDic(matchPath);

            //读取数据并以客户编号分组
            using (var file = new ExcelFileReader(dataPath))
            {
                file.ReadSheet("头程数据");
                FBAgroup = GroupByColumn(file.rows, "客户编号");

                file.ReadSheet("尾程数据");
                PACKgroup = GroupByColumn(file.rows, "客户编号");
            }

            //输出表格
            Console.WriteLine("开始分类输出");
            SortedSet<string> UIDset = new SortedSet<string>(FBAgroup.Keys);
            UIDset.UnionWith(PACKgroup.Keys);

            foreach (string cnid in UIDset)
            {
                if (!UIDMatchSever.ContainsKey(cnid))
                { continue; }

                string savePath = Path.Combine(folder, UIDMatchSever[cnid] + "客户时效表");
                if ((FBAgroup.ContainsKey(cnid)) && (PACKgroup.ContainsKey(cnid)))
                    savePath = Path.Combine(savePath, "头程+尾程客户");
                else
                    savePath = FBAgroup.ContainsKey(cnid) ? Path.Combine(savePath, "头程客户") : Path.Combine(savePath, "尾程客户");
                if (!Directory.Exists(savePath)) { Directory.CreateDirectory(savePath); }
                savePath = Path.Combine(savePath, "运德时效表" + cnid + ".xlsx");

                var sheets = new Dictionary<string, object>();
                if (FBAgroup.ContainsKey(cnid))
                    sheets.Add("头程数据", FBAgroup[cnid]);
                if (PACKgroup.ContainsKey(cnid))
                    sheets.Add("尾程数据", PACKgroup[cnid]);
                MiniExcel.SaveAs(savePath, sheets);
            }

        }



    }

}

