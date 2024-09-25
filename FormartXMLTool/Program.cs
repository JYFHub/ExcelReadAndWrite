using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Xml;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace FormartXMLTool
{
    internal class Program
    {
        private static String FileSavePath = "K:\\TotalPriceCount.xlsx";
        private static XlsxTool XSTool = new XlsxTool();
        static void Main(string[] args)
        {
            if (args.Length >= 2)
            {
                ReadXML(args[0]);
                FileSavePath = args[1];
                WriteXlsxFile();
            }
            else
            {
                ReadXML(@"K:\\TotalPriceCount.xml");
                WriteXlsxFile();
            }
            Console.WriteLine("WriteFile:Ending");
        }

        static void ReadXML(String FilePath)
        {
            XmlDocument doc = new XmlDocument();
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.IgnoreWhitespace = true;
            settings.IgnoreComments = true;
            using (XmlReader reader = XmlReader.Create(FilePath, settings))
            {
                doc.Load(FilePath);
                XmlElement Root = doc.DocumentElement;//跟元素
                foreach (XmlNode TableNode in Root.SelectNodes("Table"))
                {
                    Table tab = new Table();
                    tab.Name = TableNode.Attributes["Name"].Value;
                    String CellWidth = TableNode.Attributes["CellWidth"].Value;
                    string[] CellArray = CellWidth.Split(',');
                    for (int i = 0; i < CellArray.Length; i++)
                    {
                        tab.CellWidthList.Add(int.Parse(CellArray[i])); 
                    }
                    Row CurrentRow = null;
                    foreach (XmlNode RowNode in TableNode.SelectNodes("Row"))
                    {
                        Console.WriteLine("------------");
                        CurrentRow = new Row();
                        CurrentRow.FirstRow = int.Parse(RowNode.Attributes["FirstRow"].Value);
                        CurrentRow.LastRow = int.Parse(RowNode.Attributes["LastRow"].Value);
                        CurrentRow.CalcLen();
                        foreach (XmlNode ColNode in RowNode.ChildNodes)
                        {
                            int firstCol = int.Parse(ColNode.Attributes["FirstCol"].Value);
                            int lastCol = int.Parse(ColNode.Attributes["LastCol"].Value);
                            CurrentRow.AddCol(ColNode.InnerText,firstCol,lastCol);
                            Console.WriteLine(ColNode.InnerText);
                        }
                        tab.AddRow(ref CurrentRow);
                        Console.WriteLine("------------");
                    }
                    XSTool.AddTable(ref tab);
                } 
            }
        }

        static void WriteXlsxFile()
        {
            IWorkbook workbook = new XSSFWorkbook();  //新建IWorkbook对象
            IRow row = null;
            ICell cell = null;
            ISheet sheet = null;
            ICellStyle cellstyle = workbook.CreateCellStyle();
            cellstyle.VerticalAlignment = VerticalAlignment.Center;
            cellstyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            foreach (Table tab in XSTool.Tables)
            {
                sheet = workbook.CreateSheet(tab.Name);
                foreach (Row row1 in tab.Rows)
                {
                    row = sheet.CreateRow(row1.FirstRow);
                    //设置单元格宽度
                    for(int i = 0;i < tab.CellWidthList.Count; ++i)
                    {
                        sheet.SetColumnWidth(i,tab.CellWidthList[i]);
                    }
                    foreach (Col col1 in row1.Cols)
                    {
                        cell = row.CreateCell(col1.FirstCol);
                        cell.CellStyle = cellstyle;
                        sheet.AddMergedRegion(col1.ColMergedRange);
                        sheet.AutoSizeColumn(col1.FirstCol);
                        sheet.SetDefaultColumnStyle(col1.FirstCol, cellstyle);
                        cell.SetCellValue(col1.Value);
                    }
                }
            }
            try
            {
                using (FileStream fs = new FileStream(FileSavePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs);//保存表格文件
                }
            }
            catch (Exception)
            {
                Console.WriteLine("FileError:FileOpen!");
                MessageBox.Show("打开文件失败！请关闭打开文件之后进行重试！！！");
            }
             
        }
    }
}
