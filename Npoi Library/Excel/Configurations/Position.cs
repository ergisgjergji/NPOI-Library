using System.ComponentModel.DataAnnotations;

namespace Npoi_Library.Excel.Configurations
{
    public class Position
    {
        /// <summary>
        /// Number of row in Excel: starting from 1.
        /// </summary>
        [Required]
        public int RowIndex { get; set; }

        /// <summary>
        /// The letter of column (not case-sensitive).
        /// </summary>
        [Required]
        public string ColumnLetter { get; set; }
    }
}
