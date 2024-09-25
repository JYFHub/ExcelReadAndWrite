using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ConvertExcelToolAsset
{
    class MergedRegion
    {
        public CellRangeAddress CellRange { get; set; }
        public String CellData { get; set; }

        public MergedRegion()
        {
            CellData = "";
        }

        public void AddCellRange(CellRangeAddress range,String data)
        {
            CellRange = range;
            CellData = data;
        }

        public bool IsVaildRange(int Row,int Col)
        {
            return CellRange.IsInRange(Row, Col);
        }
    }
}
