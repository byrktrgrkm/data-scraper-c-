using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace BirimFiyatlar
{
    class BasicReq
    {
        public string UserAgent = @"Mozilla/5.0 (Windows; Windows NT 6.1) AppleWebKit/534.23 (KHTML, like Gecko) Chrome/11.0.686.3 Safari/534.23";

        private CookieContainer cJar;
        
        public string HttpGet(string url)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.CookieContainer = cJar;
                request.UserAgent = UserAgent;
                request.KeepAlive = false;
                request.Method = "GET";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream());
                return sr.ReadToEnd();
            }
            catch
            {
                return "";
            }
        }

        public string HttpPost(string url, string post, string refer = "")
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.CookieContainer = cJar;
            request.UserAgent = UserAgent;
            request.KeepAlive = false;
            request.Method = "POST";
            request.Referer = refer;

            byte[] postBytes = Encoding.ASCII.GetBytes(post);
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = postBytes.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(postBytes, 0, postBytes.Length);
            requestStream.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader sr = new StreamReader(response.GetResponseStream());

            return sr.ReadToEnd();
        }

       
        public void DebugHtml(string html)
        {
            StreamWriter sw = new StreamWriter("debug.html");
            sw.Write(html);
            sw.Close();
        }

      
        public BasicReq()
        {
            cJar = new CookieContainer();
        }

        ~BasicReq()
        {
            // Nothing here
        }
    }
}
