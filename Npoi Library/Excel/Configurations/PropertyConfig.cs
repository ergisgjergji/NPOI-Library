using NPOI.SS.UserModel;

namespace Npoi_Library.Excel.Configurations
{
    public class PropertyConfig
    {
        public int ColumnPosition { get; set; }
        public string PropertyName { get; set; }
        public string HeaderName { get; set; }
        public string DataFormat { get; set; }
        public ICellStyle CellStyle { get; set; }
    }
}
