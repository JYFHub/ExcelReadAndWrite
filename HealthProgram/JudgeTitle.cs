using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HealthProgram
{
    //判断题类型
    class JudgeTitle
    {
        //题目名称
        public String TitleName { get; set; }

        //题目专业类型
        public String TitleClassName { get; set; }

        //题目编号
        public String TitleNumName { get; set; }

        //选项A
        public String OptionAName { get; set; }

        //选项B
        public String OptionBName { get; set; }

        //正确选项
        public String RightOptionName { get; set; }

        //专家指引
        public String ExpertsGuide { get; set; }

        public JudgeTitle()
        {
            TitleName = "";
            TitleClassName = "";
            TitleNumName = "";
            OptionAName = "";
            OptionBName = "";
            RightOptionName = "";
            ExpertsGuide = "";
        }

        public bool isVaild()
        {
            if (TitleName.Length > 0 && TitleClassName.Length > 0 && OptionAName.Length > 0
                 && OptionBName.Length > 0 && RightOptionName.Length > 0)
            {
                if (Int32.Parse(TitleNumName) > 0)
                { 
                    return true;
                }
            }
            return false;
          
        }
    }
}
