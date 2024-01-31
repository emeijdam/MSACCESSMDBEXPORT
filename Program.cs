// See https://aka.ms/new-console-template for more information

using System;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Linq;
using System.Formats.Asn1;
using System.Globalization;
using CsvHelper;

string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp\bert\gemeente.mdb";
string outputfolder = @"c:\temp\bert";


{

    //DataTable tables = connection.GetSchema("Tables");

    //foreach (DataColumn column in tables.Columns)
    //{
    //    Console.WriteLine(column.ColumnName);
    //}

    //foreach (DataRow row in tables.Rows)
    //{
    //    foreach (var item in row.ItemArray)
    //    {
    //        if (row[3].ToString() == "TABLE")
    //        {
    //            Console.WriteLine("IS EEN TABEL: ", row[3].ToString());
    //        };
    //            Console.WriteLine(item.ToString());
    //    }

    //}

    dt2csv("periods");
    dt2csv("reports");
    dt2csv("ReportSyntaxes");
    dt2csv("UserCategories");
    dt2csv("UserSyntaxes");

    void dt2csv(string tablename)
    {
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            connection.Open();
            string queryString = String.Format("SELECT * FROM {0}", tablename);
            OleDbCommand command = new OleDbCommand(queryString, connection);

            OleDbDataReader reader = command.ExecuteReader();
            var dataTable = new DataTable();
            dataTable.Load(reader);

            foreach (DataRow row in dataTable.Rows)
            {
                // string res = string.Join(Environment.NewLine, row.Select(x => string.Join(" ; ", x.ItemArray)));

                foreach (var item in row.ItemArray)
                {
                    Console.WriteLine(item.ToString());

                }

            }

            using (var writer = new StreamWriter(Path.Combine(outputfolder, tablename + ".csv")))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                foreach (DataColumn dc in dataTable.Columns)
                {
                    csv.WriteField(dc.ColumnName);
                }
                csv.NextRecord();

                foreach (DataRow dr in dataTable.Rows)
                {
                    foreach (DataColumn dc in dataTable.Columns)
                    {
                        csv.WriteField(dr[dc]);
                    }
                    csv.NextRecord();
                }

                // writer.ToString().Dump();
            }
        }
    }
}