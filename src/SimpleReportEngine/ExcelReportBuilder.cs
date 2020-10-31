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

[assembly:InternalsVisibleTo("SimpleReportEngineTests")]

namespace SimpleReportEngine
{
    public static class ExcelReportBuilder 
    {
        private static readonly Regex _macro = new Regex("^\\s*%([a-zA-Z0-9.]*)%", RegexOptions.Compiled);

        public static void BuildExcelReport(this DataSet dataSet, string templateFileName, string outFileName)
        {
            if (dataSet == null)
            {
                throw new ArgumentNullException(nameof(dataSet));
            }

            if (string.IsNullOrEmpty(templateFileName))
            {
                throw new ArgumentNullException(nameof(templateFileName));
            }

            if (!File.Exists(templateFileName))
            {
                throw new ArgumentException($"Template file {templateFileName} is not found.");
            }

            if (string.IsNullOrEmpty(outFileName))
            {
                throw new ArgumentNullException(nameof(outFileName));
            }

            using var inStream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read);
            using var outStream = new FileStream(outFileName, FileMode.Create);

            BuildExcelReport(dataSet, inStream, outStream);
        }

        public static void BuildExcelReport(this DataSet dataSet, Stream inStream, Stream outStream)
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

            FillNamedAreas(dataSet, workBook);
            FillSingleCells(dataSet, workBook);

            workBook.Write(outStream);
            workBook.Close();
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
                    FillNamedArea(dataTable, areaName, workBook);
                }
            }
        }

        internal static void FillSingleCells(DataSet dataSet, XSSFWorkbook workBook)
        {
            var creationHelper = workBook.GetCreationHelper();

            foreach (ISheet sheet in workBook)
            {
                foreach (IRow row in sheet)
                {
                    foreach (ICell cell in row.Cells)
                    {
                        FillCell(dataSet, cell, creationHelper);
                    }
                }
            }
        }

        internal static void FillNamedArea(DataTable dataTable, IName areaName, XSSFWorkbook workBook)
        {
            if (dataTable == null)
            {
                return;
            }

            var creationHelper = workBook.GetCreationHelper();

            var (area, lastIndex) = GetAreaByName(workBook, areaName);
            var rowCount = dataTable.Rows.Count;

            if (rowCount == 0)
            {
                FillRange(area, creationHelper, dataTable.Columns);
                return;
            }

            if (rowCount == 1)
            {
                FillRange(area, creationHelper, dataTable.Columns, dataTable.Rows[0]);
                return;
            }

            var (nextArea, nextIndex) = (area, lastIndex);

            for (int i = 0; i < rowCount; i++)
            {
                if (i < rowCount - 1)
                {
                    (nextArea, nextIndex) = CopyArea(area, lastIndex + 1);
                }

                FillRange(area, creationHelper, dataTable.Columns, dataTable.Rows[i]);

                (area, lastIndex) = (nextArea, nextIndex);
            }
        }

        internal static (IRow[] rows, int lastIndex) GetAreaByName(XSSFWorkbook workBook, IName name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

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
            if (rows == null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            var resRows = new List<IRow>();

            for (int i = 0; i < rows.Length; i++)
            {
                var row = rows[i];

                if (row == null)
                {
                    continue;
                }

                var sheet = row.Sheet;
                var newRowIdx = startIndex + i;

                sheet.CreateRow(newRowIdx);

                var newRow = row.CopyRowTo(newRowIdx);
                resRows.Add(newRow);
            }

            return (resRows.ToArray(), startIndex + rows.Length - 1);
        }

        internal static void FillRange(
            IRow[] rows,
            ICreationHelper creationHelper,
            DataColumnCollection columns,
            DataRow dataRow = null)
        {
            IterateRange(
                rows,
                columns,
                (cell, colName) => SetCellValue(cell, dataRow?[colName], creationHelper));
        }

        internal static void IterateRange(IRow[] rows, DataColumnCollection columns, Action<ICell, string> action)
        {
            foreach (var row in rows)
            {
                if (row == null)
                {
                    continue;
                }

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
        internal static void FillCell(DataSet dataSet, ICell cell, ICreationHelper creationHelper)
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
                SetCellValue(cell, null, creationHelper);
                return;
            }

            var value = dataTable.Rows[0][colName];
            SetCellValue(cell, value, creationHelper);
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

        /// <summary>
        /// Set cell value with apropriate type
        /// </summary>
        internal static void SetCellValue(ICell cell, object value, ICreationHelper creationHelper)
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
                case Uri uri:
                    var link = creationHelper.CreateHyperlink(HyperlinkType.Url);
                    link.Address = uri.AbsoluteUri;
                    cell.Hyperlink = link;
                    cell.SetCellValue(uri.AbsoluteUri);
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }
    }
}
