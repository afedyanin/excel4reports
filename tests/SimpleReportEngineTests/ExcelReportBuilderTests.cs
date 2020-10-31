using System.IO;
using NUnit.Framework;
using SimpleReportEngine;

namespace SimpleReportEngineTests
{
    [TestFixture(Explicit = true)]
    public class ExcelReportBuilderTests
    {
        private const string _templateFileName = "Reports\\template.xlsx";
        private const string _outFileName = "Reports\\report01.xlsx";

        [Test]
        [Category("Intergation")]
        public void CanFillReport()
        {
            using var ds = DataSetFactory.CreateSimpleDataSet();
            ds.BuildExcelReport(_templateFileName, _outFileName);
            var created = File.Exists(_outFileName);
            Assert.True(created);
        }

        [TestCase("", "", "")]
        [TestCase("%col1%", "", "col1")]
        [TestCase("%table.col%", "table", "col")]
        [TestCase("%sp.m3.table.col01%", "table", "col01")]
        [TestCase("table.col01", "", "")]
        [TestCase("12344e.%cl01%aa$%", "", "")]
        public void CanGetColumnName(string source, string tableName, string colName)
        {
            var (table, column) = ExcelReportBuilder.GetColumnName(source);

            Assert.AreEqual(tableName, table);
            Assert.AreEqual(colName, column);
        }
    }
}
