using System;
using System.Net;
using System.Text.Json;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using RestSharp;
using HtmlAgilityPack;
using System.Diagnostics;

namespace WedoExpress
{
    class UPSCrawler
    {
        private RestClient m_Client;
        private Dictionary<string, string> m_Cookies = new Dictionary<string, string>()
        {
            {"ups_language_preference","en_US"},
            {"bm_sz","33688404803F23735CEF7D82959E1CC3~YAAQYpZUaFUsNSJ/AQAARQmLJA6sAcVOtRwAlO7SSnDu/rldURAiYIKXJpCz9zlxA8hKUhyca4zif2iG6FR0xrZlGcXQXK/LxoxRmgFM3PZEBYs+/jmjts2vW0R3aXBr4pO3e+RgrKRwEHmeoRlJif9KGD4c6zUooChh26/amVSyz4ZilDnwJ9ZZ574ZXDnXjcTIZbMr2aPhzi8MKW8rBMLJYPhg18LgqCsbK2vwX4DBGfq2z9I+GPKRJ+gavqpCqns66+BByI4lQ8TzKDDP/UEK9P8icYjGvc9wm1uIdeg=~3163440~3748422"},
        };
        private string dir;


        public UPSCrawler(string dir)
        {
            this.dir = dir;

            var options = new RestClientOptions("https://www.ups.com/")
            {
                UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.45 Safari/537.36"
            };

            m_Client = new RestClient(options);



            Console.WriteLine(DateTime.UtcNow.ToUniversalTime().ToString("r"));
            // Utils.CookieHeaderToDic("__ims_caf=103_70_220_82:1644372367955; com.ups.ims.lasso.sDataLassoFeb19=bdab6946a3504cf8a56ec4352281cdad:EbKQj573xVI1sxv6U7oepAwzKkV33fGherIKx6DHvrE=; sharedsession=f696b7c9-dd7f-4d92-ba2f-cb14005e6fc7:m; ups_language_preference=en_US; bm_sz=856EB3BA22BB97C0818C6BCC2B1AE196~YAAQPJZUaJ5Ugdh+AQAAHBs73A4Tb6fYJUIjazlvxhhYHPIkUTaWGb0p81LQ3tH+K099GvZonIQQD8chgGQ/TMRQ8QKfU1dLSClDhVEMEaOdLtj66AsIy2AggaxSw2hZv7/ErUBBmw7mh4crInZUKtrgLFUhVoMbysRM2SmU1NYjcWtSXcj25pdiW7RywXhIQvFSYm7Zqs0J1MM2tr1V8OsSoHJ3slkHIc9NacHBO5v7dQKKn/K8suuCMhLp95xeBJrOk6Bs7hApSf2WTFsisViuJdVZsINbseztHBMv4S8=~3360306~3421237; PIM-SESSION-ID=LBQOu2BOpKraQbxh; at_check=true; AMCVS_036784BD57A8BB277F000101@AdobeOrg=1; mbox=session#a7bc3f091f6f4a0d8f5e877170db1c6d#1644374229|PC#a7bc3f091f6f4a0d8f5e877170db1c6d.32_0#1707617170; mboxEdgeCluster=32; CONSENTMGR=consent:true|ts:1644372369353; gig_canary=false; gig_canary_ver=12833-3-27406140; _abck=654E26CAFA03E029EE6ECFC4380E5E2A~0~YAAQPJZUaKxUgdh+AQAA7yE73Adg56Hvsjjp0kZudEqD7zTg+Btoqiyii3JSMMceqPYFMN4VDq93qpcRi1xhR1QPp8QrYJo8xDi512DtEzVuV7ynKsdueR0n7pYBZ+kEEemwzrs6Fdg1aVRTpzW7zr3YuQnHU+PoqzI5HgYmCloClyC9OW3wuREGFZRJ1PWMighhIaK2xHMIQjeVoPJb6nx4ZcKqxdGF3ua1vAkdhaMT/PveloZ1t7keSUJ7ksOCXE4lwPjbN1717heMN+v5bwy2DVtbxz08TDNe4w9tkM/TDUZqzMxx2xe8JVEPG+Fq0HBjGYEJdm+B8xAmoSbIEf34WdFgKgTOfJmnms7MhqyOPsH6Pfh06KnAfO83vwIQrXxfT9oCijg5QaaNhfwbYx8q0n9+~-1~-1~-1; ak_bmsc=89EC9A2E6E7D6793301DC4B6B8B8AD29~000000000000000000000000000000~YAAQPJZUaK1Ugdh+AQAAoCM73A57J74xbssutyyqZ/AVd2xwsG812ifYg3O0i2iKl7BvqLVMzvwVoFpGyuKsIRwUw/O7dXY5u70SbqrAas12YvJ2w/kpZoGGefHqfO/zc9MfT4UKiYNZyv0nOOX3JzDNKlvf5Z57BUuGjQB3x/m3ClHUAJBI87PEGbOF1OxdAcJq64Byv9c4JFYJqjLgeTmeEJ3eRomCdXIGWJixp8oMYhuruAyePXHkoZFEcj0mR/tLq9LkkrwuZ+Z11cbWDJ+70+MQ9q48UGIY+o6PcYDNc3Vj29QSmMBG3snbSIFbnAaUIaNA4VMzgGTjEtR4DkZ14uox3m/WPy38eQzkbtatZB4QB4Hj+rCbLyh5yt8wO7XuptV/FP54NvjDAUagCJVgIC/hNUvUHRbbzDFs41dPN5gGewWfpCFT8Q5CagWe8TFVEs6uQPRuPp222uZ0yteNxhl1JzaStQmLjxrQT0ofJfosXOZTyQ==; s_vnum=1646064000642&vn=1; s_invisit=true; dayssincevisit_s=First Visit; s_cc=true; _gcl_au=1.1.1127941983.1644372371; _fbp=fb.1.1644372370735.2130584507; AMCV_036784BD57A8BB277F000101@AdobeOrg=-1124106680|MCIDTS|19033|MCMID|65027811778886630074016660026267982306|MCAAMLH-1644977170|11|MCAAMB-1644977170|RKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y|MCOPTOUT-1644379570s|NONE|MCCIDH|-1884734651|vVersion|5.2.0; aam_cms=segments=22945446|9626391|9626487|9625872|9626568|9626599|9626828|9625302; aam_uuid=65284158288785111313970232717383561254; gig_bootstrap_3_iCVSE9Ao6y9HITzXCDEN85YkhAnYbAuW1a6LOUnRKPEcwU_QCjFz7q_a1qfN5Vgd=_gigya_ver4; utag_main=v_id:017edc3b1fb90016d7ae56d8709305072007606a0086e$_sn:1$_se:2$_ss:0$_st:1644374171976$ses_id:1644372369339;exp-session$_pn:1;exp-session$fs_sample_user:false;exp-session$vapi_domain:ups.com$_prevpage:ups:us:en:lasso:login;exp-1644375972010$_prevpageid:ct1_reg_log(1ent).html;exp-1644375972012; s_nr=1644372372023-New; dayssincevisit=1644372372025");

        }

        public async Task MakeUpsPDF(string trackingNumber, string dir)
        {
            var maker = new UPSPDFMaker();

            // var info = GetTrackinginfoAsync(trackingNumber).GetAwaiter().GetResult();

            var info = await GetTrackinginfoAsync(trackingNumber);
            if (info["statusCode"].ToString() == "200")
                Console.WriteLine("获取单号成功：{0}", trackingNumber);
            else
                Console.WriteLine("获取单号失败：{0}", trackingNumber);

            // Console.WriteLine(SimpleJson.SerializeObject(info));

            //生成PDF
            var document = maker.MakeUp(info);
            document.Save(Path.Join(dir, trackingNumber) + ".pdf");
            document.Close();

            //删除签收图片
            File.Delete(info["img"].ToString());
        }

        private RestRequest LoginRequest(string username, string password)
        {

            var request = new RestRequest("lasso/login", method: Method.Post);


            request.AddHeaders(new Dictionary<string, string>(){
                {"cookie",Utils.DicToCookieHeader(m_Cookies)},
                {"accept","text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"},
                {"accept-encoding","gzip, deflate, br"},
                {"accept-language","zh-CN,zh;q=0.9"},
                {"cache-control","max-age=0"},
                {"origin","https://www.ups.com"},
                {"referer","https://www.ups.com/lasso/login?loc=en_US&returnto=https%3A%2F%2Fwww.ups.com%2Fus%2Fen%2FHome.page"},
            });

            request.AddParameter("CSRFToken", GetCsrfToken());
            request.AddParameter("loc", "en_CA");
            request.AddParameter("returnto", "https://www.ups.com/track?loc=zh_CN&Requester=lasso");
            request.AddParameter("userID", username);
            request.AddParameter("password", password);
            request.AddParameter("getTokenWithPassword", "");


            return request;

        }

        public void Login(string username, string password)
        {
            var request = LoginRequest(username, password);

            var response = m_Client.PostAsync(request).GetAwaiter().GetResult();

            if (!response.Content.Contains("Log In"))
            {
                Console.WriteLine("登录成功");
                for (int i = 0; i < response.Cookies.Count; i++)
                {
                    if (m_Cookies.ContainsKey(response.Cookies[i].Name))
                        m_Cookies[response.Cookies[i].Name] = response.Cookies[i].Value;
                    else
                        m_Cookies.Add(response.Cookies[i].Name, response.Cookies[i].Value);
                }
            }
        }

        private RestRequest TrackingRequest(string trackingNumber)
        {
            var request = new RestRequest("/track/api/Track/GetPOD", Method.Post);
            request.AddHeaders(new Dictionary<string, string>(){
                {"cookie",Utils.DicToCookieHeader(m_Cookies)},
                {"x-xsrf-token",GetXsrfToken(trackingNumber)},
                {"accept","application/json, text/plain, */*"},
                {"accept-encoding","gzip, deflate, br"},
                {"accept-language","zh-CN,zh;q=0.9"},
                {"origin","https://www.ups.com"},
                // {"referer",string.Format("https://www.ups.com/track?loc=en_US&requester=QUIC&tracknum={0}/trackdetails",trackingNumber)},
                {"user-agent","Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36"}
            });

            request.AddJsonBody(new Dictionary<string, object>(){
                {"ActCode","D"},
                {"Locale","en_CA"},
                {"TrackingNumber",new List<string>() { trackingNumber }},
            });
            return request;
        }

        public Dictionary<string, object> GetTrackinginfo(string trackingNumber)
        {
            var request = TrackingRequest(trackingNumber);
            var response = m_Client.PostAsync<Dictionary<string, object>>(request).GetAwaiter().GetResult();
            response["img"] = DownloadSignature(trackingNumber);
            return response;
        }

        public async Task<Dictionary<string, object>> GetTrackinginfoAsync(string trackingNumber)
        {
            var request = TrackingRequest(trackingNumber);
            var response = await m_Client.PostAsync<Dictionary<string, object>>(request);
            response["img"] = DownloadSignature(trackingNumber);
            return response;
        }

        private RestRequest SignatureRequest(string trackingNumber)
        {
            var request = new RestRequest("https://wwwapps.ups.com/SignatureClient/SignatureRequest");
            request.AddHeader("accept", "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8");
            request.AddHeader("accept-encoding", "gzip, deflate, br");
            request.AddHeader("accept-language", "zh-CN,zh;q=0.9");

            request.AddParameter("Requester", "TrackHTML");
            request.AddParameter("tracknum", trackingNumber);
            return request;
        }

        private string DownloadSignature(string trackingNumber)
        {
            string savePath = Path.Join(dir, trackingNumber + ".jpg");

            var request = SignatureRequest(trackingNumber);

            using (var writer = File.OpenWrite(savePath))
            {
                var response = m_Client.DownloadDataAsync(request).GetAwaiter().GetResult();
                for (int i = 0; i < response.Length; i++)
                {
                    writer.WriteByte(response[i]);
                }
            }
            return savePath;
        }

        private string GetCsrfToken()
        {
            var request = new RestRequest("/lasso/login", Method.Get);
            request.AddQueryParameter("loc", "en_CA");
            request.AddQueryParameter("returnto", "https://www.ups.com/ca/en/Home.page");
            var response = m_Client.GetAsync(request).GetAwaiter().GetResult();
            m_Client.CookieContainer.Add(response.Cookies);
            HtmlDocument document = new HtmlDocument();
            document.LoadHtml(response.Content);
            HtmlNode token = document.DocumentNode.SelectSingleNode("//input[@name='CSRFToken']");

            return token.Attributes["value"].Value;
        }

        private string GetXsrfToken(string trackingNumber)
        {
            string url = string.Format("https://www.ups.com/track?loc=en_CA&tracknum={0}&requester=ST/trackdetails", trackingNumber);
            var request = new RestRequest("/track");
            request.AddParameter("loc", "en_CA");
            request.AddParameter("tracknum", trackingNumber);
            request.AddParameter("requester", "ST/trackdetails");

            var response = m_Client.PostAsync(request).GetAwaiter().GetResult();

            for (int i = 0; i < response.Cookies.Count; i++)
            {
                var cookie = response.Cookies[i];
                if (cookie.Name == "X-XSRF-TOKEN-ST")
                {
                    m_Client.CookieContainer.Add(cookie);
                    return cookie.Value;
                }
            }

            return default(string);
        }


    }



}


