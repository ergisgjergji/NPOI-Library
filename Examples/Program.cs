
using Npoi_Library.Excel;
using Npoi_Library.Excel.Styling;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Examples
{
    class Program
    {
        static void Main(string[] args)
        {
            Test1();
            Test2();
            Test3();

            Console.WriteLine("Tests completed successfully!");
            Console.ReadLine();
        }

        private static void Test1()
        {
            string storeLocation = @"C:\Users\ergis gjergji\Desktop\Test1.xls";

            List<Model_1> data = new List<Model_1>()
            {
                new Model_1 { Prop_1 = 1, Prop_2 = "Test", Prop_3 = true, Prop_4 = DateTime.Now, Prop_5 = 1.32f },
                new Model_1 { Prop_1 = 2, Prop_2 = "Test", Prop_3 = true, Prop_4 = DateTime.Now.AddDays(-1), Prop_5 = 0.73f }
            };

            byte[] content = ExcelManager.GenerateExcel(data, new ExcelOptions 
            {
                HeaderStyle = new HeaderStyle { IsBold = false },
                BodyStyle = new BodyStyle { IsBordered = true }
            });

            using (var fileStream = File.Create(storeLocation))
            {
                fileStream.Write(content, 0, content.Length);
            }
        }

        public static void Test2()
        {
            string storeLocation = @"C:\Users\ergis gjergji\Desktop\Test2.xls";

            DataTable table = new DataTable("Test");

            DataColumn col1 = new DataColumn("Id", typeof(int));
            table.Columns.Add(col1);
            DataColumn col2 = new DataColumn("Name", typeof(string));
            table.Columns.Add(col2);
            DataColumn col3 = new DataColumn("Birthdate", typeof(DateTime));
            table.Columns.Add(col3);
            DataColumn col4 = new DataColumn("Wage", typeof(float));
            table.Columns.Add(col4);

            DataRow row = table.NewRow();
            row["Id"] = 1;
            row["Name"] = "Test";
            row["Birthdate"] = DateTime.Now.AddYears(-18);
            row["Wage"] = 550f;
            table.Rows.Add(row);

            byte[] content = ExcelManager.GenerateExcel(table, null);

            using (var fileStream = File.Create(storeLocation))
            {
                fileStream.Write(content, 0, content.Length);
            }
        }

        public static void Test3()
        {
            string templateLocation = @"C:\Users\ergis gjergji\Desktop\template.xls";
            string templateSheetName = "Sheet1";
            string storeLocation = @"C:\Users\ergis gjergji\Desktop\Test3.xls";

            List<Model_2> data = new List<Model_2>
            {
                new Model_2() { Prop_1 = 1, Prop_2 = "Test", Prop_3 = false, Prop_4 = DateTime.Now, Prop_5 = 0.35f },
                new Model_2() { Prop_1 = 2, Prop_3 = false, Prop_4 = DateTime.Now.AddDays(-3), Prop_5 = 2.12f },
            };

            Dictionary<string, Position> map1 = new Dictionary<string, Position>();
            map1.Add(nameof(Model_2.Prop_1), new Position { RowIndex = 1, ColumnLetter = "A" });
            map1.Add(nameof(Model_2.Prop_2), new Position { RowIndex = 2, ColumnLetter = "B" });
            map1.Add(nameof(Model_2.Prop_3), new Position { RowIndex = 4, ColumnLetter = "C" });
            map1.Add(nameof(Model_2.Prop_4), new Position { RowIndex = 3, ColumnLetter = "D" });
            map1.Add(nameof(Model_2.Prop_5), new Position { RowIndex = 1, ColumnLetter = "E" });

            data[0].PositionMap = map1;

            Dictionary<string, Position> map2 = new Dictionary<string, Position>();
            map2.Add(nameof(Model_2.Prop_1), new Position { RowIndex = 5, ColumnLetter = "B" });
            map2.Add(nameof(Model_2.Prop_2), new Position { RowIndex = 6, ColumnLetter = "A" });
            map2.Add(nameof(Model_2.Prop_3), new Position { RowIndex = 8, ColumnLetter = "C" });
            map2.Add(nameof(Model_2.Prop_4), new Position { RowIndex = 7, ColumnLetter = "D" });
            map2.Add(nameof(Model_2.Prop_5), new Position { RowIndex = 6, ColumnLetter = "E" });

            data[1].PositionMap = map2;

            byte[] content = ExcelManager.GenerateExcelFromTemplate(data, templateLocation, templateSheetName);

            using (var fileStream = File.Create(storeLocation))
            {
                fileStream.Write(content, 0, content.Length);
            }
        }
    }
}
