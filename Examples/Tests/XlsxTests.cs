using Examples.Models;
using Npoi_Library.Excel.Configurations;
using Npoi_Library.Excel.Styling;
using Npoi_Library.Excel.XlsManager;
using Npoi_Library.Excel.XlsxManager;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Examples.Tests
{
    public static class XlsxTests
    {
        public static void Test1()
        {
            string xlsLocation = @"C:\Users\ergis\Desktop\Test1.xlsx";
            string pdfLocation = @"C:\Users\ergis\Desktop\Test1.pdf";

            List<Customer> data = new List<Customer>()
            {
                new Customer { Id = 1, Name = "Test", Salary = 750d, BirthDate = DateTime.Now.AddYears(-18), IsActive = true },
                new Customer { Id = 2, Name = "Test", Salary = 800f, BirthDate = DateTime.Now.AddYears(-20), IsActive = false }
            };

            var eManager = new XlsxManager();

            var xlsContent = eManager.GenerateExcel(data, new ExcelOptions
            {
                HeaderStyle = new HeaderStyle { IsBold = false, IsBordered = true, FontSize = 20, FontFamily = "Arial" },
                BodyStyle = new BodyStyle { IsBordered = true, FontSize = 14 }
            });

            var pdfContent = eManager.ConvertToPdf(xlsContent);

            using (var xlsStream = File.Create(xlsLocation))
            using (var pdfStream = File.Create(pdfLocation))
            {
                xlsStream.Write(xlsContent, 0, xlsContent.Length);
                pdfStream.Write(pdfContent, 0, pdfContent.Length);
            }
        }

        public static void Test2()
        {
            string storeLocation = @"C:\Users\ergis\Desktop\Test2.xls";

            DataTable table = new DataTable("Test");

            DataColumn col1 = new DataColumn("Id", typeof(int));
            table.Columns.Add(col1);
            DataColumn col2 = new DataColumn("Name", typeof(string));
            table.Columns.Add(col2);
            DataColumn col3 = new DataColumn("Salary", typeof(double));
            table.Columns.Add(col3);
            DataColumn col4 = new DataColumn("BirthDate", typeof(DateTime));
            table.Columns.Add(col4);
            DataColumn col5 = new DataColumn("IsActive", typeof(bool));
            table.Columns.Add(col5);

            DataRow row = table.NewRow();
            row["Id"] = 1;
            row["Name"] = "Test customer";
            row["Salary"] = 750d;
            row["BirthDate"] = DateTime.Now.AddYears(-20);
            row["IsActive"] = true;
            table.Rows.Add(row);

            var eManager = new XlsxManager();
            byte[] content = eManager.GenerateExcel(table, null);

            using (var fileStream = File.Create(storeLocation))
            {
                fileStream.Write(content, 0, content.Length);
            }
        }

        public static void Test3()
        {
            string templateLocation = @"C:\Users\ergis\Desktop\template.xls";
            string templateSheetName = "Sheet1";
            string storeLocation = @"C:\Users\ergis\Desktop\Test3.xls";

            List<TemplateModel> data = new List<TemplateModel>
            {
                new TemplateModel() { Prop_1 = 1, Prop_2 = "Test", Prop_3 = false, Prop_4 = DateTime.Now, Prop_5 = 0.35f },
                new TemplateModel() { Prop_1 = 2, Prop_3 = false, Prop_4 = DateTime.Now.AddDays(-3), Prop_5 = 2.12f },
            };

            Dictionary<string, Position> map1 = new Dictionary<string, Position>();
            map1.Add(nameof(TemplateModel.Prop_1), new Position { RowIndex = 1, ColumnLetter = "A" });
            map1.Add(nameof(TemplateModel.Prop_2), new Position { RowIndex = 2, ColumnLetter = "B" });
            map1.Add(nameof(TemplateModel.Prop_3), new Position { RowIndex = 4, ColumnLetter = "C" });
            map1.Add(nameof(TemplateModel.Prop_4), new Position { RowIndex = 3, ColumnLetter = "D" });
            map1.Add(nameof(TemplateModel.Prop_5), new Position { RowIndex = 1, ColumnLetter = "E" });

            data[0].PositionMap = map1;

            Dictionary<string, Position> map2 = new Dictionary<string, Position>();
            map2.Add(nameof(TemplateModel.Prop_1), new Position { RowIndex = 5, ColumnLetter = "B" });
            map2.Add(nameof(TemplateModel.Prop_2), new Position { RowIndex = 6, ColumnLetter = "A" });
            map2.Add(nameof(TemplateModel.Prop_3), new Position { RowIndex = 8, ColumnLetter = "C" });
            map2.Add(nameof(TemplateModel.Prop_4), new Position { RowIndex = 7, ColumnLetter = "D" });
            map2.Add(nameof(TemplateModel.Prop_5), new Position { RowIndex = 6, ColumnLetter = "E" });

            data[1].PositionMap = map2;

            var eManager = new XlsxManager();
            byte[] content = eManager.GenerateExcelFromTemplate(data, templateLocation, templateSheetName);

            using (var fileStream = File.Create(storeLocation))
            {
                fileStream.Write(content, 0, content.Length);
            }
        }

        public static void Test_Read()
        {
            string location = @"C:\Users\ergis\Desktop\Test_read.xlsx";
            string sheetName = "Sheet1";

            var eManager = new XlsxManager();
            var data = eManager.ReadFromExcel<Customer>(location, sheetName, 1);
        }
    }
}
