using System;
using NUnit.Framework;
using SimpleReportEngine;

namespace SimpleReportEngineTests
{
    public class TextReportBuilderTests
    {
        [Test]
        [Category("Intergation")]
        public void CanFillReport()
        {
            using var ds = DataSetFactory.CreateSimpleDataSet();
            var res = ds.BuildTextReport();
            Assert.NotNull(res);
            Console.WriteLine(res);
        }
    }
}
