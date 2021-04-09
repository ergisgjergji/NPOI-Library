using System;
using System.Collections.Generic;
using System.Text;

namespace Npoi_Library.Excel.Configurations
{
    public class ExcelNote
    {
        public int x1 { get; set; }
        public int x2 { get; set; }
        public int y1 { get; set; }
        public int y2 { get; set; }
        public string Text { get; set; }

        public ExcelNote()
        {
        }

        public ExcelNote(int x1, int x2, int y1, int y2, string text)
        {
            this.x1 = x1;
            this.x2 = x2;
            this.y1 = y1;
            this.y2 = y2;
            Text = text;
        }
    }
}
