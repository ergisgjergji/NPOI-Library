using System;
using System.Collections.Generic;
using System.Text;

namespace Npoi_Library.Excel.Styling
{
    public class CustomCellStyle
    {
        public int FontSize { get; set; } = 11;
        public string FontFamily { get; set; } = "Calibri Light";
        public virtual bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderlined { get; set; } = false;
        public virtual bool IsBordered { get; set; } = false;
        /// <summary>
        /// RGB color.
        /// </summary>
        public virtual byte[] BackgroundColor { get; set; }
        /// <summary>
        /// RGB color.
        /// </summary>
        public byte[] FontColor { get; set; }
    }
}
