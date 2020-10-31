using System;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using SimpleReportEngine;

namespace SampleApp.Orders
{
    public static class OrdersReport
    {
        private const string _ordersDataFile = "Orders\\orders.json";
        private const string _ordersTemplateFile = "Orders\\OrdersTemplate.xlsx";
        private const string _ordersReportFile = "Orders\\OrdersReport.xlsx";

        public static string Build()
        {
            using var ds = CreateOrderDataSet(_ordersDataFile);
            ds.BuildExcelReport(_ordersTemplateFile, _ordersReportFile);
            return _ordersReportFile;
        }

        private static DataSet CreateOrderDataSet(string fileName)
        {
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

            var json = File.ReadAllText(fileName);
            var orders = JsonConvert.DeserializeObject<Order[]>(json);

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
