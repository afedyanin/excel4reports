using System;
using System.Data;

namespace DataSetReportExtensionsTests
{
    public static class DataSetFactory
    {
        public static DataSet CreateSimpleDataSet()
        {
            var ds = new DataSet();

            ds.Tables.Add("myTable01");
            ds.Tables.Add("myTable02");
            ds.Tables.Add("myTable03");

            ds.Tables["myTable01"].Columns.Add("col1", typeof(decimal));
            ds.Tables["myTable01"].Columns.Add("col2", typeof(decimal));
            ds.Tables["myTable01"].Columns.Add("col3", typeof(decimal));
            ds.Tables["myTable01"].Columns.Add("col4", typeof(decimal));
            ds.Tables["myTable01"].Columns.Add("col5", typeof(decimal));
            ds.Tables["myTable01"].Columns.Add("col6", typeof(decimal));

            ds.Tables["myTable03"].Columns.Add("m10", typeof(int));
            ds.Tables["myTable03"].Columns.Add("m11", typeof(int));
            ds.Tables["myTable03"].Columns.Add("m12", typeof(int));
            ds.Tables["myTable03"].Columns.Add("m13", typeof(int));
            ds.Tables["myTable03"].Columns.Add("m14", typeof(int));
            ds.Tables["myTable03"].Columns.Add("m15", typeof(int));

            ds.Tables["myTable02"].Columns.Add("t1");
            ds.Tables["myTable02"].Columns.Add("t2");
            ds.Tables["myTable02"].Columns.Add("t3");
            ds.Tables["myTable02"].Columns.Add("t4");

            var dr4 = ds.Tables["myTable02"].NewRow();

            dr4[0] = "column 1";
            dr4[1] = "column 2";
            dr4[2] = "column 3";
            dr4[3] = "column 4";

            ds.Tables["myTable02"].Rows.Add(dr4);

            var dr5 = ds.Tables["myTable02"].NewRow();

            dr5[0] = "column A";
            dr5[1] = "column B";
            dr5[2] = "column C";
            dr5[3] = "column D";

            ds.Tables["myTable02"].Rows.Add(dr5);

            var dr = ds.Tables["myTable01"].NewRow();

            dr[0] = 1;
            dr[1] = 2;
            dr[2] = 3;
            dr[3] = 4;
            dr[4] = 5;
            dr[5] = 6;

            ds.Tables["myTable01"].Rows.Add(dr);

            var dr2 = ds.Tables["myTable01"].NewRow();

            dr2[0] = 10;
            dr2[1] = 20;
            dr2[2] = 30;
            dr2[3] = 40;
            dr2[4] = 50;
            dr2[5] = 60;

            ds.Tables["myTable01"].Rows.Add(dr2);

            var dr3 = ds.Tables["myTable01"].NewRow();

            dr3[0] = 200.11m;
            dr3[1] = 300.45m;
            dr3[2] = 400.89m;
            dr3[3] = 800.34m;
            dr3[4] = 500.61m;
            dr3[5] = 600.42m;

            ds.Tables["myTable01"].Rows.Add(dr3);

            for (int i = 0; i < 100; i++)
            {
                dr = ds.Tables["myTable03"].NewRow();

                for (int j = 0; j < 6; j++)
                {
                    dr[j] = i + (j * 10);
                }

                ds.Tables["myTable03"].Rows.Add(dr);
            }

            return ds;
        }
    }
}
