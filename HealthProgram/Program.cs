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
namespace HealthProgram
{ 
    class Program
    {
        static List<JudgeTitle> JudgeTitleList = new List<JudgeTitle>();//判断题目
        static List<SingleChoiceTitle> TextChoiceTitleList = new List<SingleChoiceTitle>();//文字选择题目
        static List<SingleChoiceTitle> PictureChoiceTitleList = new List<SingleChoiceTitle>();//图片选择题目

        //初始化构造主函数为单元进程
        [STAThread]
        static void Main(string[] args)
        {
            String OPenFileName="";
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
                        OPenFileName = "E:/ExaminationLibrary.xlsx";
                    }
                }
                catch
                {
                    OPenFileName = "E:/ExaminationLibrary.xlsx";
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
            //清空题目选项
            JudgeTitleList.Clear();
            TextChoiceTitleList.Clear();
            PictureChoiceTitleList.Clear();
            LoadTitle(OPenFileName);//加载题目
            WriteTitleToXml(SaveFileName);//写题目到当前指定文件
        }

        static void LoadTitle(string FileName)
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
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            IRow row;           //新建当前工作表行数据
            for (int i = 0; i <= sheet.LastRowNum; i++)  //对工作表每一行
            {
                row = sheet.GetRow(i);   //row读入第i行数据
                if (row.Cells.Count() >= 3)
                {
                    if (row.GetCell(2).ToString() == "文字判断题")
                    {
                        LoadJudge(row);
                    }
                    else if (row.GetCell(2).ToString() == "文字选择题")
                    {
                        LoadSingleChoice(row, false);
                    }
                    else if (row.GetCell(2).ToString() == "图片选择题")
                    {
                        LoadSingleChoice(row,true);
                    }
                }
            }
            fileStream.Close();//关闭文件流和工作簿
            workbook.Close();
        }
        //提取判断题到数据表中
        static void LoadJudge(IRow CurrentRow)
        {
            JudgeTitle _judge = new JudgeTitle();
            _judge.TitleNumName = CurrentRow.GetCell(0).ToString();
            _judge.TitleClassName = CurrentRow.GetCell(1).ToString();
            _judge.TitleName = CurrentRow.GetCell(3).ToString();
            _judge.OptionAName = CurrentRow.GetCell(4).ToString();
            _judge.RightOptionName = CurrentRow.GetCell(4).ToString();
            _judge.OptionBName = CurrentRow.GetCell(5).ToString();
            if (CurrentRow.Cells.Count >= 7)
            {
                _judge.ExpertsGuide = CurrentRow.GetCell(8).ToString();
            }
            else
            {
                _judge.ExpertsGuide = "";
            }
            //如果题目有效,则添加进入题目数组列表
            if (_judge.isVaild())
            {
                JudgeTitleList.Add(_judge);
            }
        }
        //提取单项选择题到数据表中
        static void LoadSingleChoice(IRow CurrentRow,bool PictureChoice)
        {
            SingleChoiceTitle _siglechoice = new SingleChoiceTitle();
            _siglechoice.TitleNumName = CurrentRow.GetCell(0).ToString();
            _siglechoice.TitleClassName = CurrentRow.GetCell(1).ToString();
            _siglechoice.TitleName = CurrentRow.GetCell(3).ToString();
            _siglechoice.RightOptionName = CurrentRow.GetCell(4).ToString();
            if (CurrentRow.Cells.Count > 8)
            {
                _siglechoice.ExpertsGuide = CurrentRow.GetCell(8).ToString();
            }
            else
            {
                _siglechoice.ExpertsGuide = "";
            }
            for (int i = 4; i < 8; i++)
            {
                _siglechoice.ListOptions.Add(CurrentRow.GetCell(i).ToString());
            }
            #region //如果题目有效,则添加进入题目数组列表
                if (_siglechoice.isVaild())
                {
                    if (PictureChoice)
                    {
                        PictureChoiceTitleList.Add(_siglechoice);
                    }
                    else
                    {
                        TextChoiceTitleList.Add(_siglechoice);
                    }
                }
            #endregion
        }

        static string ConverOPenFileToSavePath(string openfile)
        {
            string temp = openfile;
            int index = temp.LastIndexOf('.');
            temp = temp.Substring(0,index);
            temp += ".xml";
            return temp;
        }

        static void WriteJudgeTitleToRoot(ref XElement root)
        {
            Console.WriteLine("WriteBeginJudgeTitle");
            XElement childs = new XElement("JudgeTitleList");
            XElement child;
            for (int i = 0; i < JudgeTitleList.Count(); ++i)
            {
                child = new XElement("JudgeTitle");
                child.SetElementValue("TitleNumName",JudgeTitleList[i].TitleNumName);
                child.SetElementValue("TitleName", JudgeTitleList[i].TitleName);
                child.SetElementValue("TitleClassName", JudgeTitleList[i].TitleClassName);
                child.SetElementValue("OptionAName", JudgeTitleList[i].OptionAName);
                child.SetElementValue("OptionBName", JudgeTitleList[i].OptionBName);
                child.SetElementValue("RightOptionName", JudgeTitleList[i].RightOptionName);
                child.SetElementValue("ExpertsGuide", JudgeTitleList[i].ExpertsGuide);
                childs.Add(child);
            }
            root.Add(childs);
            Console.WriteLine("WriteEndJudgeTitle");
        }

        static void WriteTextChoiceTitleToRoot(ref XElement root)
        {
            Console.WriteLine("WriteBeginTextChoiceTitle");
            XElement childs = new XElement("TextChoiceTitleList");
            XElement child;
            for (int i = 0; i < TextChoiceTitleList.Count(); ++i)
            {
                child = new XElement("TextChoiceTitle");
                child.SetElementValue("TitleNumName", TextChoiceTitleList[i].TitleNumName);
                child.SetElementValue("TitleName", TextChoiceTitleList[i].TitleName);
                child.SetElementValue("TitleClassName", TextChoiceTitleList[i].TitleClassName);
                child.SetElementValue("RightOptionName", TextChoiceTitleList[i].RightOptionName);
                child.SetElementValue("ExpertsGuide", TextChoiceTitleList[i].ExpertsGuide);
                XElement ListOptions = new XElement("ListOptions");
                for (int j = 0; j < TextChoiceTitleList[i].ListOptions.Count(); j++)
                {
                    string currentOpt = "ListOption";
                    currentOpt += j;
                    ListOptions.SetElementValue(currentOpt, TextChoiceTitleList[i].ListOptions[j]);
                }
                child.Add(ListOptions);
                childs.Add(child);
            }
            root.Add(childs);
            Console.WriteLine("WriteEndTextChoiceTitle");
        }

        static void WritePictureChoiceTitleToRoot(ref XElement root)
        {
            Console.WriteLine("WriteBeginPictureChoiceTitle");
            XElement childs = new XElement("PictureChoiceTitleList");
            XElement child;
            for (int i = 0; i < PictureChoiceTitleList.Count(); ++i)
            {
                child = new XElement("PictureChoiceTitle");
                child.SetElementValue("TitleNumName", PictureChoiceTitleList[i].TitleNumName);
                child.SetElementValue("TitleName", PictureChoiceTitleList[i].TitleName);
                child.SetElementValue("TitleClassName", PictureChoiceTitleList[i].TitleClassName);
                child.SetElementValue("RightOptionName", PictureChoiceTitleList[i].RightOptionName);
                child.SetElementValue("ExpertsGuide", PictureChoiceTitleList[i].ExpertsGuide);
                XElement ListOptions = new XElement("ListOptions");
                for (int j = 0; j < PictureChoiceTitleList[i].ListOptions.Count(); j++)
                {
                    string currentOpt = "ListOption";
                    currentOpt += j;
                    ListOptions.SetElementValue(currentOpt, PictureChoiceTitleList[i].ListOptions[j]);
                }
                child.Add(ListOptions);
                childs.Add(child);
            }
            root.Add(childs);
            Console.WriteLine("WriteEndPictureChoiceTitle");
        }
        static void WriteTitleToXml(string WritePath)
        {
            XDocument document = new XDocument();
            XElement root = new XElement("TitleList");
            WriteJudgeTitleToRoot(ref root);//写判断题目到题目列表中
            WriteTextChoiceTitleToRoot(ref root);//写文本选择题目到题目列表中
            WritePictureChoiceTitleToRoot(ref root);//写图片选择题目到题目列表中
            root.Save(new StreamWriter(WritePath, false, Encoding.Unicode));//保存文件到指定列表中
            Console.WriteLine("WriteSUccess");
        }
    }
}
