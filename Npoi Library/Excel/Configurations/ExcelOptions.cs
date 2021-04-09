using Npoi_Library.Excel.Styling;
using System;
using System.Collections.Generic;
using System.Text;

namespace Npoi_Library.Excel.Configurations
{
    public class ExcelOptions
    {
        public HeaderStyle HeaderStyle { get; set; } = new HeaderStyle();
        public BodyStyle BodyStyle { get; set; } = new BodyStyle();
    }
}
