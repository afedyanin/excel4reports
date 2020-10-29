using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

[assembly: InternalsVisibleTo("DataSetReportExtensionsTests")]

namespace DataSetReportExtensions
{
    public static class ExcelExtensions
    {
        private static readonly Regex _macro = new Regex("^\\s*%([a-zA-Z0-9.]*)%", RegexOptions.Compiled);

        public static void FillExcelReport(this DataSet dataSet, string inFileName, string outFileName)
        {
            if (string.IsNullOrEmpty(inFileName))
            {
                throw new ArgumentNullException(nameof(inFileName));
            }

            if (string.IsNullOrEmpty(outFileName))
            {
                throw new ArgumentNullException(nameof(outFileName));
            }

            if (!File.Exists(inFileName))
            {
                throw new ArgumentException($"Template file {inFileName} is not found.");
            }

            using var inStream = new FileStream(inFileName, FileMode.Open, FileAccess.Read);
            using var outStream = new FileStream(outFileName, FileMode.Create);

            FillExcelReport(dataSet, inStream, outStream);
        }

        public static void FillExcelReport(this DataSet dataSet, Stream inStream, Stream outStream)
        {
            if (dataSet == null)
            {
                throw new ArgumentNullException(nameof(dataSet));
            }

            if (inStream == null)
            {
                throw new ArgumentNullException(nameof(inStream));
            }

            if (outStream == null)
            {
                throw new ArgumentNullException(nameof(outStream));
            }

            var workBook = new XSSFWorkbook(inStream);
            FillWorkBook(dataSet, workBook);
            workBook.Write(outStream);
        }

        internal static void FillWorkBook(DataSet dataSet, XSSFWorkbook workBook)
        {
            FillNamedAreas(dataSet, workBook);
            FillSingleCells(dataSet, workBook);
        }

        internal static void FillNamedAreas(DataSet dataSet, XSSFWorkbook workBook)
        {
            foreach (var areaName in workBook.GetAllNames())
            {
                if (areaName.IsDeleted)
                {
                    continue;
                }

                if (dataSet.Tables.Contains(areaName.NameName))
                {
                    var dataTable = dataSet.Tables[areaName.NameName];
                    FillNamedArea(dataTable, workBook, areaName);
                }
            }
        }

        internal static void FillSingleCells(DataSet dataSet, XSSFWorkbook workBook)
        {
            foreach (ISheet sheet in workBook)
            {
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        FillCell(dataSet, cell);
                    }
                }
            }
        }

        internal static void FillNamedArea(DataTable dataTable, XSSFWorkbook workBook, IName areaName)
        {
            var (area, lastIndex) = GetAreaByName(workBook, areaName);
            var rowCount = dataTable.Rows.Count;

            if (rowCount == 0)
            {
                CleanupRange(area, dataTable.Columns);
                return;
            }

            if (rowCount == 1)
            {
                FillRange(area, dataTable.Rows[0]);
                return;
            }

            var (nextArea, nextIndex) = (area, lastIndex);

            for (int i = 0; i < rowCount; i++)
            {
                if (i < rowCount - 1)
                {
                    (nextArea, nextIndex) = CopyArea(area, lastIndex + 1);
                }

                FillRange(area, dataTable.Rows[i]);

                (area, lastIndex) = (nextArea, nextIndex);
            }
        }

        internal static (IRow[] rows, int lastIndex) GetAreaByName(XSSFWorkbook workBook, IName name)
        {
            var aref = new AreaReference(name.RefersToFormula, SpreadsheetVersion.EXCEL2007);
            var resRows = new List<IRow>();

            for (int i = aref.FirstCell.Row; i <= aref.LastCell.Row; i++)
            {
                var sheet = workBook.GetSheet(aref.FirstCell.SheetName);
                var row = sheet.GetRow(i);
                resRows.Add(row);
            }

            return (resRows.ToArray(), aref.LastCell.Row);
        }

        internal static (IRow[] rows, int lastIndex) CopyArea(IRow[] rows, int startIndex)
        {
            var resRows = new List<IRow>();

            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];
                var sheet = row.Sheet;
                var newRowIdx = startIndex + i;

                sheet.CreateRow(newRowIdx);

                var newRow = row.CopyRowTo(newRowIdx);
                resRows.Add(newRow);
            }

            return (resRows.ToArray(), startIndex + rows.Length - 1);
        }

        internal static void FillRange(IRow[] rows, DataRow dataRow)
        {
            IterateRange(
                rows,
                dataRow.Table.Columns,
                (cell, colName) => SetCellValue(cell, dataRow[colName]));
        }

        internal static void CleanupRange(IRow[] rows, DataColumnCollection columns)
        {
            IterateRange(
                rows,
                columns,
                (cell, colName) => SetCellValue(cell, null));
        }

        internal static void IterateRange(
            IRow[] rows,
            DataColumnCollection columns,
            Action<ICell, string> action)
        {
            foreach (var row in rows)
            {
                foreach (var cell in row.Cells)
                {
                    ProcessCell(cell, columns, action);
                }
            }
        }

        /// <summary>
        /// Find template macro and exec action
        /// </summary>
        internal static void ProcessCell(ICell cell, DataColumnCollection columns, Action<ICell, string> action)
        {
            if (cell.CellType != CellType.String)
            {
                return;
            }

            var cellValue = cell.StringCellValue;

            if (string.IsNullOrEmpty(cellValue))
            {
                return; 
            }

            var (_, colName) = GetColumnName(cellValue);

            if (columns.Contains(colName))
            {
                action(cell, colName);
            }
        }

        /// <summary>
        /// Find template macro and replace it with actual value
        /// </summary>
        internal static void FillCell(DataSet dataSet, ICell cell)
        {
            if (cell.CellType != CellType.String)
            {
                return;
            }

            var cellValue = cell.StringCellValue;

            if (string.IsNullOrEmpty(cellValue))
            {
                return;
            }

            var (tableName, colName) = GetColumnName(cellValue);

            if (!dataSet.Tables.Contains(tableName))
            {
                return;
            }

            var dataTable = dataSet.Tables[tableName];

            if (!dataTable.Columns.Contains(colName))
            {
                return;
            }

            if (dataTable.Rows.Count <= 0)
            {
                SetCellValue(cell, null);
                return;
            }

            var value = dataTable.Rows[0][colName];
            SetCellValue(cell, value);
        }

        /// <summary>
        /// Set cell value with apropriate type
        /// </summary>
        internal static void SetCellValue(ICell cell, object value)
        {
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            switch (value)
            {
                case bool bl:
                    cell.SetCellValue(bl);
                    break;
                case DateTime dt:
                    cell.SetCellValue(dt);
                    break;
                case double db:
                    cell.SetCellValue(db);
                    break;
                case decimal db:
                    cell.SetCellValue((double)db);
                    break;
                case long db:
                    cell.SetCellValue(db);
                    break;
                case int db:
                    cell.SetCellValue(db);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }

        /// <summary>
        /// Extract table and column name from template macro
        /// </summary>
        internal static (string tableName, string colName) GetColumnName(string cellValue)
        {
            if (string.IsNullOrEmpty(cellValue))
            {
                return (string.Empty, string.Empty);
            }

            var fullName = _macro.Match(cellValue).Groups[1]?.Value;

            if (string.IsNullOrEmpty(fullName))
            {
                return (string.Empty, string.Empty);
            }

            var items = fullName.Split('.');
            var len = items.Length;

            if (len <= 1)
            {
                return (string.Empty, items[0]);
            }

            return (items[len - 2], items[len - 1]);
        }
    }
}
