using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Excel;

namespace ToSql
{
    class Program
    {
        static void Main(string[] args)
        {
            var folder = @"C:\Users\jrupp\Desktop\TestData";
            using (var outFs = new FileStream(Path.Combine(folder, "output.sql"), FileMode.Create, FileAccess.Write))
            {
                using (var outTw = new StreamWriter(outFs))
                {
                    foreach (var file in Directory.GetFiles(folder, "*.csv"))
                    {
                        using (var fs = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Read))
                        {
                            using (var tr = new StreamReader(fs))
                            {
                                Process(outTw, file, "", new CsvReader(tr));
                            }
                        }
                    }
                    foreach (var file in Directory.GetFiles(folder, "*.xlsx"))
                    {
                        using (var wkb = new XLWorkbook(file))
                        {
                            foreach (var s in wkb.Worksheets)
                            {
                                // ignore hidden worksheets
                                if (s.Visibility != XLWorksheetVisibility.Visible)
                                    continue;
                                Process(outTw, file, s.Name, new CsvReader(new ExcelParser(s)));
                            }
                        }
                    }
                }
            }
        }

        private static void Process(TextWriter tw, string file, string sheet, CsvReader reader)
        {
            var tableName = (Path.GetFileNameWithoutExtension(file) + "_" + sheet).TrimEnd('_');
            reader.ReadHeader();
            tw.WriteLine($"declare @{tableName} table (");
            tw.WriteLine(string.Join("," + Environment.NewLine, reader.FieldHeaders.Select(h => $"  [{h.Trim().Replace("?", "")}] nvarchar(max)")));
            tw.WriteLine(");");

            tw.WriteLine("BEGIN");
            while (reader.Read())
            {
                tw.Write($"insert into @{tableName} values (");
                tw.Write(string.Join(", ", Enumerable.Range(0, reader.FieldHeaders.Length).Select(h => "N'" + reader.GetField<string>(h).Replace("'", "''").Trim() + "'")));
                tw.WriteLine(");");
            }
            tw.WriteLine("END");
        }
    }
}
