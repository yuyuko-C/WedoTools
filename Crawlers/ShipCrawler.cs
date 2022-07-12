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
using static System.Console;


namespace WedoExpress
{
    class ShipCrawler
    {
        private RestClient m_Client;
        private CookieCollection m_Cookies;
        private string m_CookieHeader;


        public ShipCrawler(string json)
        {
            var options = new RestClientOptions("https://www.shipxy.com/")
            {
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36 Edg/99.0.1150.55"
            };
            m_Client = new RestClient(options);
            var request = new RestRequest("Ship/Index");
            var response = m_Client.GetAsync(request).GetAwaiter().GetResult();

            m_Cookies = new CookieCollection();
            List<string> cookies = new List<string>();
            for (int i = 0; i < response.Cookies.Count; i++)
            {
                cookies.Add(response.Cookies[i].Name + "=" + response.Cookies[i].Value);
            }
            m_CookieHeader = string.Join("; ", cookies);
            WriteLine(m_CookieHeader);

            request = new RestRequest("ship/GetShip");
            request.AddHeader("cookie", m_CookieHeader);
            request.AddJsonBody<Dictionary<string, string>>(new Dictionary<string, string>(new Dictionary<string, string>(){
                {"mmsi","353155000"}
            }));

            WriteLine(m_Client.PostAsync(request).GetAwaiter().GetResult().Content);

        }



    }


}