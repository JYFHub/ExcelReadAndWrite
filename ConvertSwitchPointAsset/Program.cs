using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace ConvertSwitchPointAsset
{
    class Program
    {
        static string GB2312ToUTF8(string str)
        {
            Encoding uft8 = Encoding.GetEncoding("utf-8");
            Encoding gb2312 = Encoding.GetEncoding("gb2312");
            byte[] temp = uft8.GetBytes(str);
            byte[] temp1 = Encoding.Convert(gb2312,uft8, temp);
            string result = uft8.GetString(temp1);
            return result;
        }
        static void ReadExcelData(string FileName)
        {
            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(FileName,FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            catch
            {
                Console.WriteLine("FileError:Open File Faild");
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            String FileHead = "FileHead:" + sheet.LastRowNum;
            Console.WriteLine(FileHead);

            IRow row;           //新建当前工作表行数据
            String FileRow = "FileRow:";
            String CurrentRowData = "";
            for (int i = 1; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                CurrentRowData = "";
                if (row!= null && row.LastCellNum >= 9)
                {
                    for (int j = 0; j < 9; j++)  //对工作表每一列
                    {
                        if(row.GetCell(j) != null)
                        {
                            CurrentRowData += row.GetCell(j).ToString(); //获取i行j列数据
                        }
                        CurrentRowData += ",";
                    }
                    CurrentRowData = CurrentRowData.Replace("\n","+");
                    Console.WriteLine(FileRow + CurrentRowData);
                } 
            }
            fileStream.Close();
            workbook.Close();
            
        }
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            //String path = "C:/Users/JYF/Desktop/CM96X.xlsx";
           // ReadExcelData(path);//读取Excel表格数据文件
            ReadExcelData(args[0]);//读取Excel表格数据文件

        }
    }
}
