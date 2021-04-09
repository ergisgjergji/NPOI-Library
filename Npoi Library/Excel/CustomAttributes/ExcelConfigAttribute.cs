using System;

namespace Npoi_Library.Excel.CustomAttributes
{
    public class ExcelConfig : Attribute
    {
        public int ColumnPosition { get; set; } = 0;
        public string HeaderName { get; set; }
        public string DataFormat { get; set; }
    }
}
