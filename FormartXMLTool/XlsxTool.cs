using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FormartXMLTool
{
    internal class MathTools
    {
        public static int MAX(int p1, int p2)
        {
            return p1 > p2 ? p1 : p2;
        }
        public static int Min(int p1, int p2)
        {
            return p1 < p2 ? p1 : p2;
        }
        public static int MAX(int p1,int p2,int p3)
        {
            int temp =  p1 > p2 ? p1 : p2;
            temp = temp > p3 ? temp : p3;
            return temp;
        }
    }
    
    internal class Col
    {
        public String Value { get; set; }
        public int FirstCol { get; set; }
        public int LastCol { get; set; }

        private int collen = 0;
        public int ColLen
        {
            get
            {
                return FirstCol + collen;
            }
            set
            {
                collen = value;
            }
        }
        public CellRangeAddress ColMergedRange { get; set; }
        public Col()
        {
            Value = "";
            FirstCol = 0;
            LastCol = 0;
        }
        public void CalcLen()
        {
            ColLen = LastCol - FirstCol;
            if(ColLen < 0)
            {
                ColLen = 0;
            }
        }
        public Col(String value, ref CellRangeAddress ColCellRange)
        {
            Value = value;
            ColMergedRange = ColCellRange;
            FirstCol = ColCellRange.FirstColumn;
            LastCol = ColCellRange.LastColumn;
        }
    }
    internal class Row
    {
        public List<Col> Cols { get; set; }
        public int MaxColLen { get; set; }

        private int rowlen = 0;
        public int RowLen { 
            get 
            {
                return FirstRow + rowlen;
            }
            set 
            {
                rowlen = value;
            } 
        }
        public int FirstRow { get; set; }
        public int LastRow { get; set; }
        public void CalcLen() {
            RowLen = LastRow - FirstRow;
        }
        public Row()
        {
            Cols = new List<Col>();
            Cols.Clear();
            FirstRow = 0;
            LastRow = 0;
        }
        public void AddCol(String value,int FirstCol,int LastCol)
        {
            CellRangeAddress cellRange = new CellRangeAddress(FirstRow, LastRow, FirstCol, LastCol);
            Col col = new Col(value, ref cellRange);
            col.CalcLen();
            Cols.Add(col);
            MaxColLen = MathTools.MAX(MaxColLen,col.ColLen);
        }
    }
    internal class Table {
        public String Name { get; set; }

        public int MaxRowLen { get; set; }
        public int MaxColLen { get; set; }

        public List<Row> Rows { get; set; }

        public List<int> CellWidthList { get; set; }
        public Table()
        {
            Rows = new List<Row>();
            Rows.Clear();
            Name = "";
            MaxRowLen = 0;
            MaxColLen = 0;
            CellWidthList = new List<int>();
        }

        public void AddRow(ref Row currentRow)
        {
            Rows.Add(currentRow);
            MaxRowLen = MathTools.MAX(currentRow.RowLen, MaxRowLen);
            MaxColLen = MathTools.MAX(currentRow.MaxColLen, MaxColLen);
        }
    }
    internal class XlsxTool
    {
        public List<Table> Tables { get; set; }

        public XlsxTool()
        {
            Tables = new List<Table>();
            Tables.Clear();
        }

        public void AddTable(ref Table currentTable)
        {
            Tables.Add(currentTable);
        }
    }
}
