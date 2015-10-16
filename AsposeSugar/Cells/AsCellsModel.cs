using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AsposeSugar
{
    /// <summary>
    /// excel列信息
    /// </summary>
    public class ExcelColumns
    {
        public string field { get; set; }
        public string cellsAlign { get; set; }
        public string align { get; set; }
        public string text { get; set; }
        public string columngroup { get; set; }
    }
    /// <summary>
    /// 表头分组信息
    /// </summary>
    public class ExcelColumnsGroup
    {
        public string text { get; set; }
        public string align { get; set; }
        public string name { get; set; }
    }

    public class AsposeCellInfo
    {
        public string text { get; set; }
        public int x { get; set; }
        public int xCount { get; set; }
        public int y { get; set; }
        public int yCount { get; set; }
        public string groupName { get; set; }
    }
}
