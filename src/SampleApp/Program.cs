using System;
using System.Diagnostics;
using SampleApp.Books;
using SampleApp.Orders;

namespace SampleApp
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            if (args == null)
            {
                throw new ArgumentNullException(nameof(args));
            }

            if (args.Length < 2)
            {
                throw new ArgumentException("Please, specify data source (books, orders) and report format (excel, text)");
            }

            var isBooks = string.Compare(args[0], "books", StringComparison.OrdinalIgnoreCase) == 0;
            var useExcel = string.Compare(args[1], "excel", StringComparison.OrdinalIgnoreCase) == 0;

            var file = isBooks ?
                useExcel ? BooksReport.BuildExcel() : BooksReport.BuildText() :
                useExcel ? OrdersReport.BuildExcel() : OrdersReport.BuildText();

            OpenFile(file);
        }

        private static void OpenFile(string fileName)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = fileName,
                UseShellExecute = true,
            };

            Process.Start(psi);
        }
    }
}
