using Npoi_Library.Excel.CustomAttributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace Examples.Models
{
    public class Customer
    {
        [ExcelConfig(ColumnPosition = 1)]
        public int Id { get; set; }

        [ExcelConfig(ColumnPosition = 2)]
        public string Name { get; set; }

        [ExcelConfig(ColumnPosition = 3)]
        public double Salary { get; set; }

        [ExcelConfig(ColumnPosition = 4)]
        public DateTime? BirthDate { get; set; }

        [ExcelConfig(ColumnPosition = 5)]
        public bool IsActive { get; set; }
    }
}
