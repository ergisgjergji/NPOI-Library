namespace Npoi_Library.Excel.Styling
{
    public class HeaderStyle : CustomCellStyle
    {
        public override bool IsBold { get; set; } = true;
        public override bool IsBordered { get; set; } = true;
        public override byte[] BackgroundColor { get; set; } = new byte[] { 187, 255, 184 };
    }
}
