using System;
using System.Collections.Generic;
using System.Text;

namespace Npoi_Library.Excel.CustomAttributes.Configurations
{
    public class ExcelTemplateDataSection
    {
        public int x1 { get; set; }
        public int x2 { get; set; }
        public int y1 { get; set; }
        public int y2 { get; set; }

        public ExcelTemplateDataSection()
        {
        }

        public ExcelTemplateDataSection(int x1, int x2, int y1, int y2)
        {
            this.x1 = x1;
            this.x2 = x2;
            this.y1 = y1;
            this.y2 = y2;
        }
    }
}
