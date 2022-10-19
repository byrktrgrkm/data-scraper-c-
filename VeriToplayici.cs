using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using Newtonsoft.Json;



namespace BirimFiyatlar
{
    class VeriToplayici
    {
        private const string baseUrl = "https://www.birimfiyat.com";
        private const string url = "/poz-arama/{book}?page={page}&pageSize={pageSize}&SearchText=";
        private int page = 1,
            pageSize = 500,
            maxPage = 0,
            maxSize = 0;


        private int currentPage = 0;
        private int currentSize = 0;

        private string book = "çsb";

        private bool start = false , _stoped = false;

        public int limit = 0;

        private bool _step10k = false;

        private bool step1Ok {
            get
            {
                return this._step10k;
            }
            set
            {
                if (this.step1Ok == false && value == true )
                {
                    this.step2();
                }

                this._step10k = value;
            }

        }

        private List<string> urls;

        public BasicReq basicReq;

        private Main mainForm;


        public VeriToplayici(Main form)
        {
            basicReq = new BasicReq();
            urls = new List<string>();
            mainForm = form;

        }
        public void Start()
        {
        

            this.start = true;
            this._stoped = false;
            

            if (this.step1Ok)
            {
                
                this.step2();
           

                return;
            }

            this._step10k = false;

            if (this.currentPage != 0)
            {
                this.page = this.currentPage;
            }

            this.get(this.page);

        }
        public void Stop()
        {
            this.start = false;
          
        }

        public void get(int page)
        {
            if (!this.start) return;

            if (this.step1Ok) return;

            Thread thread = new Thread(t =>
                {
                    this.setPage(page);
                    mainForm.setInfo("Sayfa "+page.ToString()+ " / " + this.maxPage + " verisi toplanıyor..");
                    string html = basicReq.HttpGet(this.getUrl());
                    this.parse(html);

                    this.currentPage = page;

                    if (!this.start)
                    {
                        this.stoped();
                    
                        return;
                    }
                    // Eğer gerekli linkler toplanmışsa diğer aşamaya geç!
    

                    if (this.page < this.maxPage)
                    {
                        this.get(this.page + 1);
                    }
                    else
                    {
                        step1Ok = true;
                    }

                    
                })
            {IsBackground = true };
            

            thread.Start();
            
        }

        private void stoped()
        {
            if (!this.start &&  !this._stoped)
            {
                this.mainForm.setInfo("Durduruldu.");
                this.mainForm.Stoped();
            }

            this._stoped = true;

        }


        public void step2()
        {
            Thread thread = new Thread(t =>
            { 
            for (int i = this.currentSize; i < this.urls.Count; i++)
            {
                if (!start)
                {
                    this.stoped();
                    return;
                }

                string item = this.urls[i];
                

                this.mainForm.setOrderNo(i + 1);
                string link = VeriToplayici.baseUrl + System.Net.WebUtility.HtmlDecode(item);
                this.mainForm.setInfo(System.Net.WebUtility.HtmlDecode(item));
                string html = basicReq.HttpGet(link);


                this.detailParse(html, item);

                this.currentSize = i + 1;

                int value = this.limit > 0 ? (int)((i + 1) / Convert.ToDouble(this.urls.Count) * 100) :  (int)Math.Ceiling(this.urls.Count / Convert.ToDouble(this.maxSize));

                if (value > 0)
                    this.mainForm.loadingBar(value );
                
            }

            //tamamlandı.

            this.mainForm.Completed();

            })
            { IsBackground = true };


            thread.Start();
        }

        public void parse(string html)
        {
            try
            {
                HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(html);

                if (this.maxPage == 0)
                {
                    this.calculateMaxPage(htmlDocument);
                }

                HtmlNodeCollection htmlNodes = htmlDocument.DocumentNode.SelectNodes("//div[@class='workitem-link-btn']");

                if (htmlNodes != null)
                {
                    for (int i = 0; i < htmlNodes.Count; i++)
                    {
                        if (this.step1Ok)
                        {
                            break;
                        }

                        HtmlNode node = htmlNodes[i];

                        // son olarak gelen elemanları sırası ile burada okuyacağız.
                        HtmlAgilityPack.HtmlDocument _subDocument = new HtmlAgilityPack.HtmlDocument();
                        _subDocument.LoadHtml(node.InnerHtml);

                        HtmlNode linkNode = _subDocument.DocumentNode.SelectSingleNode("//a");

                        string detayLink = linkNode.Attributes["href"].Value;

                        this.urls.Add(detayLink);

                        mainForm.setInfo(detayLink);

                        this.mainForm.setSelectedLink(this.urls.Count);

                        if (this.limit > 0 && this.limit == this.urls.Count)
                        {
                            this.step1Ok = true;
                        }

                        if (this.maxSize > 0 && this.step1Ok == false)
                        {
                            int value = (int)Math.Ceiling(this.urls.Count / Convert.ToDouble(this.maxSize));
                            if (value > 0)
                                this.mainForm.loadingBar(value);
                        }

                    }
                }
            }
            catch
            {
                this.mainForm.setInfo("Sayfa bağlantısı sağlanamadı.");
            }
        }
        private void calculateMaxPage(HtmlDocument htmlDocument)
        {
            try
            {
                HtmlNode item = htmlDocument.DocumentNode.SelectSingleNode("//div[@class='result-count']");
                string totalSting = item.InnerText;

                string pattern = @"\d+";
                RegexOptions options = RegexOptions.Multiline;

                Match m = Regex.Match(totalSting, pattern, options);

                int total = Convert.ToInt32(m.Value);
                double result = total / Convert.ToDouble(this.pageSize);

                this.maxPage = (int)Math.Ceiling(result);

                this.maxSize = total;

                this.mainForm.setTotalLink(total);

               
            }
            catch
            {
                this.maxPage = 0;
            }
        }
        

        public void detailParse(string html, string url)
        {
            try
            {
                HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                htmlDocument.LoadHtml(html);

                HtmlNode htmlNode = htmlDocument.DocumentNode.SelectSingleNode("//*[@id='detail-view']/div[1]/div/div/table");

                HtmlNodeCollection list = htmlNode.SelectNodes("//td");

                //poz no
                string pozNo = System.Net.WebUtility.HtmlDecode(list[0].InnerText.Trim());

                pozNo = Regex.Replace(pozNo, @"\s+", " ");

                string tanim = System.Net.WebUtility.HtmlDecode(list[1].InnerText.Trim());

                string birim = System.Net.WebUtility.HtmlDecode(list[2].InnerText.Trim());

                string kurum = System.Net.WebUtility.HtmlDecode(list[3].InnerText.Trim());

                string fasikul = System.Net.WebUtility.HtmlDecode(list[4].InnerText.Trim());


                string pattern = @"let data = (\[.*\]);";

                RegexOptions options = RegexOptions.Multiline;

                Match m = Regex.Match(html, pattern, options);
                List<PriceItem> data = new List<PriceItem>();
                if (m.Groups[1].Value != "[]")
                {
                    data = JsonConvert.DeserializeObject<List<PriceItem>>(m.Groups[1].Value);
                }



                this.mainForm.addTable(pozNo, tanim, birim, kurum, fasikul, data);


            }
            catch
            {
                this.mainForm.setInfo($"{url} hata ile karşılaştı.");
            }

        }


        public string getUrl()
        {
            return VeriToplayici.baseUrl + VeriToplayici.url
                .Replace("{page}", this.page.ToString())
                .Replace("{book}", this.book)
                .Replace("{pageSize}", this.pageSize.ToString());
        }

        public void setPage(int page)
        {
            if(page < 1)
            {
                this.page = 1;
            }
            else
            {
                this.page = page;
            }
        }
        public void setBook(string book)
        {
            this.book = book;
        }
        public void setPageSize(int pageSize)
        {
            if (pageSize < 1)
            {
                this.pageSize = 500;
            }
            else
            {
                this.pageSize = pageSize;
            }
        }

        public void setLimit(int limit)
        {
            this.limit = limit;
        }
        public void refresh()
        {
            this.step1Ok = false;
            this.currentSize = 0;
            this.currentPage = 0;
            this.maxPage = 0;
            this.maxSize = 0;
            this.setPage(1);
            this.urls.Clear();
        }

    }


    public class PriceItem
    {
        public int Year { get; set; }
        public string UnitPrice { get; set; }
    }

}
