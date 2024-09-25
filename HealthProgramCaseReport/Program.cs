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
using System.Windows.Forms.ComponentModel;

namespace HealthProgramCaseReport
{
    class Program
    {
        static List<ChildrenMapping> ChildrenMappingList = new List<ChildrenMapping>();//初始化静态保存选项列表
        static CaseReportTemple patientReportTemple = new CaseReportTemple();
        //初始化构造主函数为单元进程
        [STAThread]
        static void Main(string[] args)
        {
            String OPenFileName = "";
            String SaveFileName = "";
            #region 指定打开或者关闭文件路径
            if (args.Count() <= 0)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
                ofd.ValidateNames = true;
                ofd.CheckPathExists = true;
                ofd.CheckFileExists = true;
                ofd.Title = "打开Excel题库文件";
                ofd.InitialDirectory = "E:\\";
                ofd.Multiselect = false;
                ofd.AddExtension = true;
                try
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        OPenFileName = ofd.FileName;
                    }
                    else
                    {
                        OPenFileName = "E:/病例01.xlsx";
                    }
                }
                catch
                {
                    OPenFileName = "E:/病例01.xlsx";
                }
                //转化打开路径到保存文件路径
                SaveFileName = ConverOPenFileToSavePath(OPenFileName);
            }
            else
            {
                OPenFileName = args[0];
                SaveFileName = args[1];

            }
            #endregion
            ChildrenMappingList.Clear();//清空保存对象列表
            LoadReportData(OPenFileName);//加载报表数据
            WriteTitleToXml(SaveFileName);//保存报表数据到XMl
        }
        //转换提取文件路径到病例文件默认打开路径
        static string ConverOPenFileToSavePath(string openfile)
        {
            string temp = openfile;
            int index = temp.LastIndexOf('.');
            temp = temp.Substring(0, index);
            temp += ".xml";
            return temp;
        }

        //加载病例报告数据
        static void LoadReportData(string FileName)
        {
            if (!FileName.EndsWith(".xls") && !FileName.EndsWith(".xlsx"))
            {
                MessageBox.Show("打开的数据表格格式不支持");
                return;
            }
            XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作

            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
            }
            catch
            {
                MessageBox.Show("请关闭打开的Excel文件");
                return;
            }
            ReadReportDataSheet(workbook.GetSheetAt(0));
            ReadDataMappingSheet(workbook.GetSheetAt(1));
            fileStream.Close();//关闭文件流和工作簿
            workbook.Close();
        }
        static void ReadReportDataSheet(ISheet sheet)
        {
            if (sheet.LastRowNum <= 0 || sheet.LastRowNum != 16)
                return;
            if (sheet.GetRow(0).Cells.Count != 2 || sheet.GetRow(0).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientName = "";
            }
            else
            {
                patientReportTemple.patientName = sheet.GetRow(0).Cells[1].ToString();
            }
            if (sheet.GetRow(1).Cells.Count != 2 || sheet.GetRow(1).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientAge = "";
            }
            else
            {
                patientReportTemple.patientAge = sheet.GetRow(1).Cells[1].ToString();
            }
            if (sheet.GetRow(2).Cells.Count != 2 || sheet.GetRow(2).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientBirthday = "";
            }
            else
            {
                patientReportTemple.patientBirthday = sheet.GetRow(2).Cells[1].ToString();
            }
            if (sheet.GetRow(3).Cells.Count != 2 || sheet.GetRow(3).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientIllnessNum = "";
            }
            else
            {
                patientReportTemple.patientIllnessNum = sheet.GetRow(3).Cells[1].ToString();
            }
            if (sheet.GetRow(4).Cells.Count != 2 || sheet.GetRow(4).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientConptionNum = "";
            }
            else
            {
                patientReportTemple.patientConptionNum = sheet.GetRow(4).Cells[1].ToString();
            }
            if (sheet.GetRow(5).Cells.Count != 2 || sheet.GetRow(5).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientProductionNum = "";
            }
            else
            {
                patientReportTemple.patientProductionNum = sheet.GetRow(5).Cells[1].ToString();
            }
            if (sheet.GetRow(6).Cells.Count != 2 || sheet.GetRow(6).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientLastMenstrualPeriod = "";
            }
            else
            {
                patientReportTemple.patientLastMenstrualPeriod = sheet.GetRow(6).Cells[1].ToString();
            }
            if (sheet.GetRow(7).Cells.Count != 2 || sheet.GetRow(7).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientContraception = "";
            }
            else
            {
                patientReportTemple.patientContraception = sheet.GetRow(7).Cells[1].ToString();
            }
            if (sheet.GetRow(8).Cells.Count != 2 || sheet.GetRow(8).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientCategory = "";
            }
            else
            {
                patientReportTemple.patientCategory = sheet.GetRow(8).Cells[1].ToString();
            }
            if (sheet.GetRow(9).Cells.Count != 2 || sheet.GetRow(9).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientCheckAgainst = "";
            }
            else
            {
                patientReportTemple.patientCheckAgainst = sheet.GetRow(9).Cells[1].ToString();
            }
            if (sheet.GetRow(10).Cells.Count != 2 || sheet.GetRow(10).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientTBS = "";
            }
            else
            {
                patientReportTemple.patientTBS = sheet.GetRow(10).Cells[1].ToString();
            }
            if (sheet.GetRow(11).Cells.Count != 2 || sheet.GetRow(11).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientLeuCorrheaRegular = "";
            }
            else
            {
                patientReportTemple.patientLeuCorrheaRegular = sheet.GetRow(11).Cells[1].ToString();
            }
            if (sheet.GetRow(12).Cells.Count != 2 || sheet.GetRow(12).Cells[1].ToString().Length <= 0)
            {
                patientReportTemple.patientHPV = "";
            }
            else
            {
                patientReportTemple.patientHPV = sheet.GetRow(12).Cells[1].ToString();
            }
            if (sheet.GetRow(13).Cells.Count >= 0)
            {
                for (int i = 1; i < sheet.GetRow(13).Cells.Count; i++)
                {
                    patientReportTemple.patientVaginalImageList.Add(sheet.GetRow(13).Cells[i].ToString());
                }
            }
            if (sheet.GetRow(14).Cells.Count >= 0)
            {
                for (int i = 1; i < sheet.GetRow(14).Cells.Count; i++)
                {
                    patientReportTemple.patientLesionLocationList.Add(sheet.GetRow(14).Cells[i].ToString());
                }
            }
            if (sheet.GetRow(15).Cells.Count >= 0)
            {
                for (int i = 1; i < sheet.GetRow(15).Cells.Count; i++)
                {
                    patientReportTemple.patientEvaluatesImpressionsList.Add(sheet.GetRow(15).Cells[i].ToString());
                }
            }
            if (sheet.GetRow(16).Cells.Count >= 0)
            {
                for (int i = 1; i < sheet.GetRow(16).Cells.Count; i++)
                {
                    patientReportTemple.patientFurtherHandlingOfCommentsList.Add(sheet.GetRow(16).Cells[i].ToString());
                }
            }
        }
        static void ReadDataMappingSheet(ISheet sheet)
        { 
            IRow row;
            ChildrenMapping _childMapping;
            for (int i = 0; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                if (!(row != null && row.Cells.Count > 0 && row.GetCell(0).ToString().Length > 0))
                {
                    continue;
                }
                _childMapping = new ChildrenMapping();
                _childMapping.ChilderenOPtionName = row.GetCell(0).ToString();
                for (int j = 1; j < row.Cells.Count; ++j)
                {
                    if (row.GetCell(j).ToString().Length > 0)
                    { 
                        _childMapping.ChildrenOPtionList.Add(row.GetCell(j).ToString()); 
                    }
                }
                if (_childMapping.isVaild())
                {
                    ChildrenMappingList.Add(_childMapping);
                }
            }
        }
        static void WriteTitleToXml(string WritePath)
        {
            XDocument document = new XDocument();
            XElement root = new XElement("ReportData");
            WriteReportDataToRoot(ref root);//保存报告数据到XML中
            WriteDataMappingToRoot(ref root);//保存报告映射数据到XML中
            root.Save(new StreamWriter(WritePath, false, Encoding.Unicode));//保存文件到指定列表中
            Console.WriteLine("WriteSUccess");
        }
        static void WriteReportDataToRoot(ref XElement root)
        {
            Console.WriteLine("WriteBeginpatientReport");
            XElement patientReport = new XElement("patientReport");
            patientReport.SetElementValue("patientName", patientReportTemple.patientName);
            patientReport.SetElementValue("patientAge", patientReportTemple.patientAge);
            patientReport.SetElementValue("patientBirthday", patientReportTemple.patientBirthday);
            patientReport.SetElementValue("patientIllnessNum", patientReportTemple.patientIllnessNum);
            patientReport.SetElementValue("patientConptionNum", patientReportTemple.patientConptionNum);
            patientReport.SetElementValue("patientProductionNum", patientReportTemple.patientProductionNum);
            patientReport.SetElementValue("patientLastMenstrualPeriod", patientReportTemple.patientLastMenstrualPeriod);
            patientReport.SetElementValue("patientContraception", patientReportTemple.patientContraception);
            patientReport.SetElementValue("patientCategory", patientReportTemple.patientCategory);
            patientReport.SetElementValue("patientCheckAgainst", patientReportTemple.patientCheckAgainst);
            patientReport.SetElementValue("patientTBS", patientReportTemple.patientTBS);
            patientReport.SetElementValue("patientLeuCorrheaRegular", patientReportTemple.patientLeuCorrheaRegular);
            patientReport.SetElementValue("patientHPV", patientReportTemple.patientHPV);
            XElement patientVaginalImageList = new XElement("patientVaginalImageList");
            for (int i = 0; i < patientReportTemple.patientVaginalImageList.Count; i++)
            {
                 string currentOpt = "ListOption";
                 currentOpt += i;
                 patientVaginalImageList.SetElementValue(currentOpt, patientReportTemple.patientVaginalImageList[i]);
            }
            XElement patientLesionLocationList = new XElement("patientLesionLocationList");
            for (int i = 0; i < patientReportTemple.patientLesionLocationList.Count; i++)
            {
                string currentOpt = "ListOption";
                currentOpt += i;
                patientLesionLocationList.SetElementValue(currentOpt, patientReportTemple.patientLesionLocationList[i]);
            }
            XElement patientEvaluatesImpressionsList = new XElement("patientEvaluatesImpressionsList");
            for (int i = 0; i < patientReportTemple.patientEvaluatesImpressionsList.Count; i++)
            {
                string currentOpt = "ListOption";
                currentOpt += i;
                patientEvaluatesImpressionsList.SetElementValue(currentOpt, patientReportTemple.patientEvaluatesImpressionsList[i]);
            }
            XElement patientFurtherHandlingOfCommentsList = new XElement("patientFurtherHandlingOfCommentsList");
            for (int i = 0; i < patientReportTemple.patientFurtherHandlingOfCommentsList.Count; i++)
            {
                string currentOpt = "ListOption";
                currentOpt += i;
                patientFurtherHandlingOfCommentsList.SetElementValue(currentOpt, patientReportTemple.patientFurtherHandlingOfCommentsList[i]);
            }
            patientReport.Add(patientVaginalImageList);
            patientReport.Add(patientLesionLocationList);
            patientReport.Add(patientEvaluatesImpressionsList);
            patientReport.Add(patientFurtherHandlingOfCommentsList);
            root.Add(patientReport);
            Console.WriteLine("WriteEndpatientReport");
        }
        static void WriteDataMappingToRoot(ref XElement root)
        {
            Console.WriteLine("WriteBeginpatientReport");
            XElement glossaryTable = new XElement("glossaryTable");
            XElement child;
            for (int i = 0; i < ChildrenMappingList.Count(); ++i)
            {
                string currentkey = "key";
                currentkey += i;
                child = new XElement(currentkey);
                child.SetAttributeValue("TitleName", ChildrenMappingList[i].ChilderenOPtionName);              
                for (int j = 0; j < ChildrenMappingList[i].ChildrenOPtionList.Count; j++)
                {
                    string currentvalue = "Value";
                    currentvalue += j;
                    child.SetElementValue(currentvalue, ChildrenMappingList[i].ChildrenOPtionList[j]); 
                }
                glossaryTable.Add(child);
            }
            root.Add(glossaryTable);
            Console.WriteLine("WriteEndpatientReport");
        }
    }  
}
