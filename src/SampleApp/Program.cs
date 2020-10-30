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
        private const string _booksDataFile = "Data\\books.json";
        private const string _booksTemplateFile = "Reports\\BooksTemplate.xlsx";
        private const string _booksReportFile = "Reports\\BooksReport.xlsx";

        private const string _ordersDataFile = "Data\\orders.json";
        private const string _ordersTemplateFile = "Reports\\OrdersTemplate.xlsx";
        private const string _ordersReportFile = "Reports\\OrdersReport.xlsx";

        public static void Main(string[] args)
        {
            //// GenerateBooksReport();
            GenerateOrdersReport();
        }

        private static void GenerateBooksReport()
        {
            using var ds = CreateBookDataSet(_booksDataFile);
            ds.FillExcelReport(_booksTemplateFile, _booksReportFile);

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = _booksReportFile,
                UseShellExecute = true,
            };

            Process.Start(psi);
        }

        private static void GenerateOrdersReport()
        {
            using var ds = CreateOrderDataSet(_ordersDataFile);
            ds.FillExcelReport(_ordersTemplateFile, _ordersReportFile);

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = _ordersReportFile,
                UseShellExecute = true,
            };

            Process.Start(psi);
        }

        private static DataSet CreateBookDataSet(string fileName)
        {
            var json = File.ReadAllText(fileName);
            var books = JsonConvert.DeserializeObject<Book[]>(json);

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
            bt.Columns.Add("Website", typeof(Uri));

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
                    new Uri(b.Website));
            }

            return ds;
        }

        private static DataSet CreateOrderDataSet(string fileName)
        {
            var json = File.ReadAllText(fileName);
            var orders = JsonConvert.DeserializeObject<Order[]>(json);

            var ds = new DataSet();

            var ht = ds.Tables.Add("Header");

            ht.Columns.Add("ReportTitle", typeof(string));
            ht.Columns.Add("ReportDate", typeof(DateTime));

            ht.Rows.Add("Sales report", DateTime.Now.ToString());

            var bt = ds.Tables.Add("myOrders");

            bt.Columns.Add("OrderDate", typeof(DateTime));
            bt.Columns.Add("Region", typeof(string));
            bt.Columns.Add("Rep", typeof(string));
            bt.Columns.Add("Item", typeof(string));
            bt.Columns.Add("Units", typeof(int));
            bt.Columns.Add("UnitCost", typeof(decimal));
            bt.Columns.Add("Total", typeof(decimal));

            foreach (var o in orders)
            {
                bt.Rows.Add(
                    o.OrderDate,
                    o.Region,
                    o.Rep,
                    o.Item,
                    o.Units,
                    o.UnitCost,
                    o.Total);
            }

            return ds;
        }
    }
}
