using System.Collections.Generic;

namespace Npoi_Library.Excel.Configurations
{
    public interface IPositionable
    {
        Dictionary<string, Position> PositionMap { get; set; }
    }
}
