using System.Diagnostics;
using SampleApp.Books;
using SampleApp.Orders;

namespace SampleApp
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            var booksReport = BooksReport.Build();
            OpenFile(booksReport);

            var ordersReport = OrdersReport.Build();
            OpenFile(ordersReport);
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
