using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using ConsoleTables;

namespace SimpleReportEngine
{
    public static class TextReportBuilder
    {
        public static string BuildTextReport(this DataSet dataSet)
        {
            if (dataSet == null)
            {
                throw new ArgumentNullException(nameof(dataSet));
            }

            var sb = new StringBuilder();

            foreach (DataTable table in dataSet.Tables)
            {
                sb.AppendLine(BuildTextReport(table));
            }

            return sb.ToString();
        }

        public static string BuildTextReport(this DataTable dataTable)
        {
            if (dataTable == null)
            {
                throw new ArgumentNullException(nameof(dataTable));
            }

            var cto = new ConsoleTableOptions
            {
                EnableCount = false,
                Columns = GetColumnNames(dataTable),
            };

            var ct = new ConsoleTable(cto);

            foreach (DataRow row in dataTable.Rows)
            {
                ct.AddRow(row.ItemArray);
            }

            return ct.ToString();
        }

        private static IEnumerable<string> GetColumnNames(DataTable dataTable)
        {
            var res = new List<string>();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                res.Add(dataTable.Columns[i].ColumnName);
            }

            return res;
        }
    }
}
