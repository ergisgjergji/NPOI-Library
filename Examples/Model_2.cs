using Npoi_Library.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Examples
{
    public class Model_2 : IPositionable
    {
        public int Prop_1 { get; set; }
        public string Prop_2 { get; set; }
        public bool Prop_3 { get; set; }
        public DateTime Prop_4 { get; set; }
        public float Prop_5 { get; set; }
        public Dictionary<string, Position> PositionMap { get; set; }
    }
}
