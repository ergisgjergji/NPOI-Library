using Npoi_Library.Excel.CustomAttributes;
using System;

namespace Examples
{
    public class Model_1
    {
        [ExcelConfig(ColumnPosition = 1, HeaderName = "Id")]
        public int Prop_1 { get; set; }
        public string Prop_2 { get; set; }
        public bool Prop_3 { get; set; }
        [ExcelConfig(ColumnPosition = 2, HeaderName = "Timestamp", DataFormat = "dd-mm-yyyy")]
        public DateTime Prop_4 { get; set; }
        public float Prop_5 { get; set; }
    }
}
