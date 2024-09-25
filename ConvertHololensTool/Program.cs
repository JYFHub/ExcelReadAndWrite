using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConvertHololensTool
{
    class Program
    {
        static List<String> Lines = new List<string>();
        static void ReadExcelData(string FileName,string FileSavePath)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            catch
            {
                Console.WriteLine("FileError:Open File Faild");
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            String FileHead = "表格总行数:" + sheet.LastRowNum;
            Console.WriteLine(FileHead);

            IRow row;           //新建当前工作表行数据
            String CurrentRowData = "";
            String CurrentID = "";
            String ID = "";
            for (int i = 7; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                CurrentRowData = "";
                CurrentID = "";
                ID = "";
                Console.WriteLine("-----------------------------");
                if (row != null)
                {
                    for (int j = 0; j < 6; j++)  //对工作表每一列
                    {
                        ICell CurrentCell = row.GetCell(j);
                        if (CurrentCell != null)
                        {
                            CurrentRowData = CurrentCell.ToString().Trim(); //获取i行j列数据
                            if (j == 0)
                            {
                                ID = CurrentRowData;
                                CurrentID = CurrentRowData + "+";
                            }
                            else if (j == 2)
                            {
                                XSSFRichTextString XSSTextString = (XSSFRichTextString)CurrentCell.RichStringCellValue;
                                if (XSSTextString.NumFormattingRuns >= 2)
                                {
                                    int length = XSSTextString.NumFormattingRuns;
                                    int index = XSSTextString.GetIndexOfFormattingRun(length-1);
                                    CurrentID += CurrentRowData.Substring(index);
                                    CurrentID += "+MS";  
                                }
                            }
                            if (j == 5 && CurrentID.Length > 2 && ID.Length > 0)
                            {
                                CurrentCell.SetCellValue(CurrentID);
                                Lines.Add(CurrentID);
                            }
                        }
                        else
                        {
                            if (j == 5 && CurrentID.Length > 2 && ID.Length > 0)
                            {
                                CurrentCell = row.CreateCell(j);
                                CurrentCell.SetCellValue(CurrentID);
                                Lines.Add(CurrentID);
                            }
                        }
                        
                        Console.WriteLine("RowData=" + CurrentRowData + "   CurrentID = " + CurrentID);

                    }
                    Console.WriteLine("-----------------------------");
                }
            }
            sheet.AutoSizeColumn(5);//自动宽度
            try
            {
                using (FileStream fs = new FileStream(FileSavePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fs);//保存表格文件
                }
            }
            catch (Exception)
            {
                Console.WriteLine("文件写入错误!");
            }
            fileStream.Close();
            workbook.Close();

        }
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            String Readpath = @"I:\Program\项目功能制作需求\2023\西安局项目\C70检修实训项目\C70检车员作业流程.xlsx";
            String Savepath = @"I:\Program\项目功能制作需求\2023\西安局项目\C70检修实训项目\C70检车员作业流程1.xlsx";
            ReadExcelData(Readpath, Savepath);//读取Excel表格数据文件

            System.IO.File.WriteAllLines(@"F:\Project\UE4\UE427\XianRailwayAR\trunk\C70Services\Lines.txt",Lines);
            Console.ReadKey();
        }
    }
}
