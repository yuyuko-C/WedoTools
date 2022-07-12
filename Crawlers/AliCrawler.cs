using System;
using System.Net;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using RestSharp;
using RestSharp.Extensions;
using HtmlAgilityPack;
using System.Diagnostics;



namespace WedoExpress
{

    class AliCrawler
    {
        private RestClient m_Client;
        private CookieCollection m_Cookies;
        private string m_CookieHeader;
        private int failCount = 0;

        public AliCrawler(string json)
        {
            Dictionary<string, string> headers = new Dictionary<string, string>(){
            {"User-Agent","Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/94.0.4577.82 Safari/537.36"}
            };

            var options = new RestClientOptions("https://www.alibaba.com/")
            {
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36"
            };
            m_Client = new RestClient(options);
            m_Client.AddDefaultHeaders(headers);

            m_Cookies = new CookieCollection();
            var listJsonDic = JsonSerializer.Deserialize<List<Dictionary<string, JsonElement>>>(json);
            foreach (var jsonDic in listJsonDic)
            {
                Cookie cookie = new Cookie()
                {
                    Version = 0,
                    Name = jsonDic["name"].ToString(),
                    Value = jsonDic["value"].ToString(),
                    Port = null,
                    Domain = jsonDic["domain"].ToString(),
                    Path = jsonDic["path"].ToString(),
                    Secure = jsonDic["secure"].GetBoolean(),
                    Expires = DateTime.MaxValue,
                    Discard = true,
                    Comment = null,
                    CommentUri = null,
                    HttpOnly = jsonDic["httpOnly"].GetBoolean(),

                };
                m_Cookies.Add(cookie);
            }

            List<string> cookies = new List<string>();
            for (int i = 0; i < m_Cookies.Count; i++)
            {
                cookies.Add(m_Cookies[i].Name + "=" + m_Cookies[i].Value);
            }
            m_CookieHeader = string.Join("; ", cookies);
        }


        public string SearchKeyword(string keyword, int page = 1, string indexArea = "company_en")
        {
            keyword = keyword.Trim().Replace(" ", "_");
            return string.Format("trade/search?keyword={0}&indexArea={1}&page={2}", keyword, indexArea, page);
        }

        public (List<string> contacts, string next) ParseSearchPage(string pageUrl)
        {
            // Console.WriteLine("SearchPage:" + pageUrl);
            List<string> contacts = new List<string>();

            // pageUrl = "trade/search?viewType=L&n=38&indexArea=company_en&page=39&f1=y&keyword=massage_armchair";

            var request = new RestRequest(pageUrl);
            var response = m_Client.GetAsync(request).GetAwaiter().GetResult();
            var a = response.Cookies[0];


            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(response.Content);
            var nodes = document.DocumentNode.SelectNodes(@"//h2[@class='title ellipsis']//a[@target='_blank']");

            string next = null;
            if (nodes != null)
            {
                foreach (var item in nodes)
                {
                    contacts.Add(item.Attributes["href"].Value.Replace("company_profile.html#top-nav-bar", "contactinfo.html"));
                }
                HtmlNode node = document.DocumentNode.SelectSingleNode(@"//a[@class='next']");
                if (node == null)
                {
                    next = null;
                    Console.WriteLine("can't find next");
                }
                else
                {
                    next = node.Attributes["href"].Value.TrimStart("www.alibaba.com/".ToCharArray());
                }
            }

            return (contacts, next);
        }

        public (HtmlNodeCollection contactTable, string encryptAccountId) ReadContactPage(string pageUrl)
        {
            var request = new RestRequest(pageUrl);
            var response = m_Client.GetAsync(request).GetAwaiter().GetResult();
            string content = System.Web.HttpUtility.UrlDecode(response.Content);

            // 异常检测容错
            if (content.Contains("Oops! We can't find the page you're looking for."))
                return (null, null);
            else
            {
                HtmlNodeCollection contactTable = GetCompanyInfoTable(content);
                string encryptAccountId = GetEncryptAccountId(content);

                if ((contactTable == null) || (encryptAccountId == null))
                {
                    failCount++;
                    if (failCount >= 3)
                        throw new Exception("Fail too many times. pageUrl:" + pageUrl);
                    return ReadContactPage(pageUrl);
                }
                else
                    failCount = 0;
                return (contactTable, encryptAccountId);
            }
        }

        public Dictionary<string, string> ParseContactPage(string pageUrl, HtmlNodeCollection contactTable, string encryptAccountId)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            //读取公司信息
            foreach (var row in contactTable)
            {
                string rowtext = row.InnerText.Replace("\r\n", "").Replace("\t", "").Replace(" ", "").Replace("`", "");
                int splitIndex = rowtext.IndexOf(':');
                dic.Add(rowtext.Substring(0, splitIndex), rowtext.Substring(splitIndex + 1));
            }
            //读取联系方式
            Dictionary<string, string> contactDic = GetContactInfo(pageUrl, encryptAccountId);
            foreach (var item in contactDic)
            {
                dic.Add(item.Key, item.Value);
            }

            return dic;
        }

        private HtmlNodeCollection GetCompanyInfoTable(string content)
        {

            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(content);
            var table = document.DocumentNode.SelectNodes("//table[@class='contact-table']/tr[@class='info-item']");
            if (table == null)
            {
                table = document.DocumentNode.SelectNodes("//table[@class='company-info-data table']/tr[not(contains(@class,'hide'))]");
            }
            return table;
        }

        private string GetEncryptAccountId(string PageSource)
        {
            string encryptAccountId = Regex.Match(PageSource, "\"encryptAccountId\":\"(.+?)\"").Groups[1].Value;
            if (encryptAccountId == "")
            {
                encryptAccountId = Regex.Match(PageSource, "data-account-id=\"(.+?)\"").Groups[1].Value;
                // Console.WriteLine("encryptAccountId：" + encryptAccountId);
            }
            return encryptAccountId == "" ? null : encryptAccountId;
        }

        private string GetCtoken()
        {
            string ctoken = null;
            var Cookies = m_Client.CookieContainer.GetCookies(m_Client.BuildUri(new RestRequest()));
            for (int i = 0; i < Cookies.Count; i++)
            {
                var item = Cookies[i];
                if (item.Name == "xman_us_t")
                    if (item.Value.Contains("ctoken"))
                    {
                        foreach (var p in item.Value.Split("&"))
                        {
                            if (p.Contains("ctoken"))
                            {
                                ctoken = p.Trim("ctoken=".ToCharArray());
                                Console.WriteLine("ctoken" + ctoken);
                                break;
                            }
                        }
                    }
            }

            return ctoken;
        }

        private Dictionary<string, string> GetContactInfo(string pageUrl, string encryptAccountId)
        {
            UriBuilder tel_api = new UriBuilder(pageUrl);
            tel_api.Path = "/event/app/contactPerson/showContactInfo.htm";

            var request = new RestRequest(tel_api.ToString());
            request.AddQueryParameter("encryptAccountId", encryptAccountId);
            request.AddQueryParameter("ctoken", GetCtoken());
            request.AddHeader("cookie", m_CookieHeader);

            var response = m_Client.GetAsync(request).GetAwaiter().GetResult();
            // Console.WriteLine(response.Content);

            var jsonDic = JsonSerializer.Deserialize<Dictionary<string, object>>(response.Content);
            var contactDic = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonDic["contactInfo"].ToString());

            return contactDic;
        }

    }


}