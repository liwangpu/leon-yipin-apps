using System;

namespace EpplusHelper
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        public string Tile { get; set; }
        public int Column { get; set; }
        public ExcelColumnAttribute(string title, int column = 1)
        {
            Tile = title;
            Column = column;
        }
    }
}
