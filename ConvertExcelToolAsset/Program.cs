using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace ConvertExcelToolAsset
{
    class Program
    {
        public static List<MergedRegion> CellRegionList = new List<MergedRegion>();
        static String GetDataInRegion(int Row,int Col)
        {
            String Value = "";
            foreach (var ItemRegion in CellRegionList)
            {
                if(ItemRegion.IsVaildRange(Row,Col))
                {
                    Value = ItemRegion.CellData;
                    return Value;
                }
            }
            return Value;
        }
        static void ReadExeclSheet(ISheet sheet)
        {
            //获取并且遍历合并的范围单元格
            for (int i = 2; i < sheet.NumMergedRegions; ++i)
            {
                CellRangeAddress range = sheet.GetMergedRegion(i);
                MergedRegion Region = new MergedRegion();
                Region.AddCellRange(range, sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn).ToString());
                CellRegionList.Add(Region);
            }

            String FileHead = "FileHead:" + sheet.LastRowNum;
            Console.WriteLine(FileHead);
            IRow row;           //新建当前工作表行数据
            String FileRow = "FileRow:";
            String CurrentRowData = "";
            for (int i = 2; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                CurrentRowData = "";
                if (row != null)
                {
                    bool bSkipSendData = false;
                    for (int j = 0; j < row.LastCellNum; j++)  //对工作表每一列
                    {
                        if (row.GetCell(j) != null)
                        {
                            String RegionData = GetDataInRegion(i,j);
                            if(RegionData != "")
                            {
                                CurrentRowData += RegionData;
                            }
                            else
                            {
                                RegionData = row.GetCell(j).ToString(); //获取i行j列数据
                                CurrentRowData += RegionData;
                            }
                            if (j <= 2 && RegionData.Length <= 0)
                            {
                                bSkipSendData = true;
                            }
                        }
                        else
                        {
                            if (j <= 2)
                            {
                                bSkipSendData = true;
                            }
                        }
                        CurrentRowData += ",";
                    }
                   
                    if (!bSkipSendData)
                    {
                        CurrentRowData = CurrentRowData.Replace("\n", "###");
                        Console.WriteLine(FileRow + CurrentRowData);
                    }
                }
            } 
        }
        static void CreateReadExcel(string FileName)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            catch
            {
                Console.WriteLine("FileError:Open File Faild");
                return;
            }
           
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            ReadExeclSheet(sheet);
            fileStream.Close();
            workbook.Close();

        } 
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            CellRegionList.Clear();//清空读取数据
            //String path = @"C:/Users/JYF/Desktop/模板.xlsx";
            //CreateReadExcel(path);//读取Excel表格数据文件
            CreateReadExcel(args[0]);//读取Excel表格数据文件
            //Console.ReadLine();
        }
    }
}
