using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel.Syndication;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ScrapeAssignment
{
    class Program
    {
        static void Main(string[] args)
        {
            var program = new Program();
            Console.WriteLine("Enter RSS Feed link:");

            string url = Console.ReadLine();    //"http://allafrica.com/tools/headlines/rdf/africa/headlines.rdf"; 
                                                //"http://allafrica.com/tools/headlines/rdf/business/headlines.rdf";
            var feedData = program.GetFeedData(url);
            Console.WriteLine("Enter Sample Data File Location(xlsx):");
            var fileName = Console.ReadLine();  // @"D:\rev_rest_africa.xlsx";
            program.ProcessExcel(fileName, feedData);
        }

        public List<RssFeedData> GetFeedData(string rssLink)
        {
            Console.WriteLine("Started receiving RSS Feed");
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Parse;
            settings.ValidationType = ValidationType.DTD;
            XmlReader reader = XmlReader.Create(rssLink, settings);
            SyndicationFeed feed = SyndicationFeed.Load(reader);
            reader.Close();
            var feedList = new List<RssFeedData>();
            foreach (SyndicationItem item in feed.Items)
            {
                feedList.Add(new RssFeedData
                {
                    Summary = item.Summary.Text,
                    Title = item.Title.Text,
                    Link = item.Id
                });
            }
            Console.WriteLine("Finished receiving RSS Feed");
            return feedList;
        }

        public void ProcessExcel(string fileName, List<RssFeedData> rssFeedData = null)
        {
            Console.WriteLine("Started processing RSS feed with given data");
            using (ExcelPackage package = new ExcelPackage(new FileInfo(fileName)))
            {
                var workSheet = package.Workbook.Worksheets.FirstOrDefault();

                int maxColumn = workSheet.Dimension.End.Column;
                int maxRow = workSheet.Dimension.End.Row;
                int minColumn = workSheet.Dimension.Start.Column;
                int minRow = workSheet.Dimension.Start.Row;

                var companyNames = new Dictionary<int, string>();

                for (int i = minRow + 1; i <= maxRow; i++)
                {
                    companyNames.Add(i, workSheet.Cells[i, 1].Value.ToString());
                }
                var updated = false;
                foreach (var companyName in companyNames)
                {
                    var _companyName = companyName.Value;
                    int _maxColumn = maxColumn;
                    foreach (var feedData in rssFeedData)
                    {
                        if (feedData.Summary.IndexOf(_companyName) != -1 || feedData.Title.IndexOf(_companyName) != -1)
                        {
                            updated = true;
                            workSheet.Cells[companyName.Key, _maxColumn + 1].Value = feedData.Link;
                        }
                    }
                }
                Console.WriteLine(updated ? "Data updated" : "No data found");
                package.Save();
            }
            Console.WriteLine("Finished processing RSS feed with given data");
        }
    }

    public class RssFeedData
    {
        public string Summary { get; set; }
        public string Title { get; set; }
        public string Link { get; set; }
    }
}
