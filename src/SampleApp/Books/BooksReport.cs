using System;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using SimpleReportEngine;

namespace SampleApp.Books
{
    public static class BooksReport
    {
        private const string _booksDataFile = "Books\\books.json";
        private const string _booksTemplateFile = "Books\\BooksTemplate.xlsx";
        private const string _booksReportFile = "Books\\BooksReport.xlsx";
        private const string _booksTextFile = "Books\\BooksReport.txt";

        public static string BuildExcel()
        {
            using var ds = CreateBookDataSet(_booksDataFile);
            ds.BuildExcelReport(_booksTemplateFile, _booksReportFile);
            return _booksReportFile;
        }

        public static string BuildText()
        {
            using var ds = CreateBookDataSet(_booksDataFile);
            var text = ds.Tables["myBooks"].BuildTextReport();
            File.WriteAllText(_booksTextFile, text);
            return _booksTextFile;
        }

        private static DataSet CreateBookDataSet(string fileName)
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
            bt.Columns.Add("Website", typeof(Uri));

            var json = File.ReadAllText(fileName);
            var books = JsonConvert.DeserializeObject<Book[]>(json);

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
    }
}
