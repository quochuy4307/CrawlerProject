using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Web;
using System.Net;
using System.IO;
using OfficeOpenXml;
using System.Runtime.ConstrainedExecution;

namespace CrawlerProject
{
    internal class Program
    {
        static void Main(string[] args)
        {
            HtmlWeb htmlWeb = new HtmlWeb()
            {
                AutoDetectEncoding = true,
                OverrideEncoding = Encoding.UTF8, //Hiển thị tiếng việt
            };
            var url = "https://www.toyota.com.vn/danh-sach-xe";
            //HtmlDocument load trang web
            HtmlDocument htmlDocument = htmlWeb.Load(url);
            var cars = new List<Car>();
            var divs = htmlDocument.DocumentNode.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("col-12 col-md-6 col-lg-4 product-show-item")).ToList();
            //crawl dữ liệu
            foreach ( var div in divs ) {
                var car = new Car
                {
                    Model = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("product-item-content")).FirstOrDefault()
                    .Descendants("h2").FirstOrDefault().InnerText,
                    Price = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("product-item-content")).FirstOrDefault()
                    .Descendants("span").Skip(1).FirstOrDefault().InnerText,
                    Currency = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("product-item-content")).FirstOrDefault()
                    .Descendants("span").Skip(2).FirstOrDefault().InnerText,
                    Link = "https://www.toyota.com.vn" + div.Descendants("a").FirstOrDefault().ChildAttributes("href").FirstOrDefault().Value,
                    ImageUrl = div.Descendants("img").FirstOrDefault().ChildAttributes("src").FirstOrDefault().Value,
                    NumberOfSeats = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").FirstOrDefault().InnerText,
                    Designs = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").Skip(1).FirstOrDefault().InnerText,
                    FuelType = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").Skip(2).FirstOrDefault().InnerText,
                    Origin = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").Skip(3).FirstOrDefault().InnerText,
                    Features = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").Skip(4).FirstOrDefault().InnerText,
                    Engine = div.Descendants("div").Where(x => x.GetAttributeValue("Class", "").Equals("content-product-description")).FirstOrDefault()
                    .Descendants("p").Skip(5).FirstOrDefault().InnerText,
                };
                car.Features = WebUtility.HtmlDecode(car.Features);
                car.Engine = WebUtility.HtmlDecode(car.Engine);
                car.Origin = WebUtility.HtmlDecode(car.Origin);
                car.Designs = WebUtility.HtmlDecode(car.Designs);
                cars.Add(car);
            }
            //xuất ra excel
            // If you are a commercial business and have
            // purchased commercial licenses use the static property
            // LicenseContext of the ExcelPackage class:
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"D:\ExportExcelCrawlData\Sample.xlsx");
            using(ExcelPackage excel = new ExcelPackage(file))
            {
                ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets["sheet1"];
                excelWorksheet.Cells.LoadFromCollection(cars, true);
                FileInfo excelFile = new FileInfo(@"D:\ExportExcelCrawlData\Result.xlsx");
                excel.SaveAs(excelFile);
            }
            var imgs = htmlDocument.DocumentNode.Descendants("img").ToList();
            WebClient client = new WebClient();
            int i = 0;
            var urlWeb = @"https://www.toyota.com.vn/";
            foreach (var item in imgs)
            {
                if (item.Attributes["data-src"] == null)
                {
                    string urlDownload = WebUtility.HtmlDecode(item.Attributes["src"].Value);
                    try
                    {
                        if(!urlDownload.Contains(urlWeb))
                            client.DownloadFile(urlWeb + urlDownload, @"D:\ExportExcelCrawlData\" + i + ".jpg");
                        else
                            client.DownloadFile(urlDownload, @"D:\ExportExcelCrawlData\" + i + ".jpg");
                    }
                    catch
                    {

                    }
                }
                else
                {
                    string urlDownload = WebUtility.HtmlDecode(item.Attributes["data-src"].Value);
                    try
                    {
                        if(!urlDownload.Contains(urlWeb))
                            client.DownloadFile(urlWeb + urlDownload, @"D:\ExportExcelCrawlData\" + i + ".jpg");
                        else
                            client.DownloadFile(urlDownload, @"D:\ExportExcelCrawlData\" + i + ".jpg");
                    }
                    catch
                    {

                    }
                }
                i++;
            }

            Console.ReadLine();
        }
    }
}
