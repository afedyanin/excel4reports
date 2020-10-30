using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using DataSetReportExtensions;
using Newtonsoft.Json;
using SampleApp.Model;

namespace SampleApp
{
    public static class Program
    {
        private const string _dataFile = "Data\\books.json";
        private const string _templateFile = "Reports\\BookTemplate.xlsx";
        private const string _reportFile = "Reports\\Report.xlsx";

        public static void Main(string[] args)
        {
            var json = File.ReadAllText(_dataFile);
            var books = JsonConvert.DeserializeObject<Book[]>(json);
            using var ds = CreateDataSet(books);
            ds.FillExcelReport(_templateFile, _reportFile);

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = _reportFile,
                UseShellExecute = true,
            };

            Process.Start(psi);
        }

        private static DataSet CreateDataSet(Book[] books)
        {
            var ds = new DataSet();

            var ht = ds.Tables.Add("Header");
            ht.Columns.Add("ReportTitle", typeof(string));
            ht.Columns.Add("ReportDate", typeof(DateTime));
            ht.Rows.Add("My favorite books", DateTime.Now.ToString());

            var bt = ds.Tables.Add("myBooks");
            bt.Columns.Add("Isbn", typeof(long));
            bt.Columns.Add("Title", typeof(string));
            bt.Columns.Add("Subtitle", typeof(string));
            bt.Columns.Add("Author", typeof(string));
            bt.Columns.Add("Published", typeof(DateTime));
            bt.Columns.Add("Publisher", typeof(string));
            bt.Columns.Add("Pages", typeof(int));
            bt.Columns.Add("Description", typeof(string));
            bt.Columns.Add("Website", typeof(string));

            foreach (var b in books)
            {
                bt.Rows.Add(
                    b.Isbn,
                    b.Title,
                    b.Subtitle,
                    b.Author,
                    b.Published,
                    b.Publisher,
                    b.Pages,
                    b.Description,
                    b.Website);
            }

            return ds;
        }
    }
}
