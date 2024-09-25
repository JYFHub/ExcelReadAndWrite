using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Xml.Linq;
using System.Windows.Forms;

namespace ExcelReadAndWrite
{
    class Program
    {
        static List<CheckType> StateList = new List<CheckType>();
        static void ReadExcelData(string FileName)
        {
            
            XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作

            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream=null;
            try
            { 
                fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            catch
            {
                MessageBox.Show("打开文件失败！请检查文件是否关闭或者被加密！");
                return;
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            IRow row;           //新建当前工作表行数据
            for (int i = 0; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                CheckType _type = new CheckType();
                if (row != null)
                {
                    if(row.LastCellNum >= 5)
                    {
                        for (int j = 0; j < 5; j++)  //对工作表每一列
                        {
                            string cellValue = "";
                            try
                            {
                                cellValue = row.GetCell(j).ToString(); //获取i行j列数据
                                switch (j)
                                {
                                    case 0:
                                        _type.Num = int.Parse(cellValue);
                                        break;
                                    case 1:
                                        _type.LableName = cellValue;
                                        break;
                                    case 2:
                                        _type.DataPortName = cellValue;
                                        break;
                                    case 3:
                                        _type.WriteOperatorStateToList(cellValue);
                                        break;
                                    case 4:
                                        if (cellValue.Equals("Y"))
                                        { 
                                            _type.UsePressLight = true;
                                        }
                                        break;
                                }
                            }
                            catch
                            {
                                
                            } 
                        }
                        if (_type.IsValid())
                        {
                            StateList.Add(_type);
                        }
                    }
                }
            }
            fileStream.Close();
            workbook.Close();
        }
        static void WriteDataToXml(string WritePath)
        {
            string Directorypath =  Path.GetDirectoryName(WritePath);
            if(!Directory.Exists(Directorypath))
            {
                Directory.CreateDirectory(Directorypath);
            }
            XDocument document = new XDocument();
            XElement root = new XElement("CheckActors");
            XElement child;
            for (int i = 0; i < StateList.Count(); ++i)
            {
                child = new XElement("CheckActor");
                child.SetElementValue("Num", StateList[i].Num);
                child.SetElementValue("LableName", StateList[i].LableName);
                child.SetElementValue("DataPortName", StateList[i].DataPortName);
                child.SetElementValue("PressLightState", StateList[i].UsePressLight);
                XElement operatorStats =  new XElement("OPeratorStats");
                XElement operatorStat;
                for (int j = 0; j < StateList[i].OperatorStateList.Count(); ++j)
                { 
                    operatorStat = new XElement("OPeratorStat");
                    operatorStat.SetElementValue("StatNum", StateList[i].OperatorStateList[j].Num);
                    operatorStat.SetElementValue("StatName", StateList[i].OperatorStateList[j].Value);
                    operatorStats.Add(operatorStat);
                }
                child.Add(operatorStats);
                root.Add(child);
            }
            root.Save(new StreamWriter(WritePath, false, Encoding.Unicode));
            Console.WriteLine("WriteSUccess");
        }
        static void Main(string[] args)
        {
            StateList.Clear();
            //ReadExcelData("C:/Users/JYF/Desktop/检查2007.xlsx");
           // WriteDataToXml("J:/UE425/Program/Scene_U1899_1801/Plugins/CheckFillActorPlugin/ThirdParty/CheckResult/检查2007.xml");

            ReadExcelData(args[0]);//读取Excel表格数据文件
            WriteDataToXml(args[1]);
        }
    }
}
