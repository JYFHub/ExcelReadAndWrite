using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HealthProgram
{
    class SingleChoiceTitle
    {
        //题目名称
        public String TitleName { get; set; }

        //题目专业类型
        public String TitleClassName { get; set; }

        //选项列表
        public List<String> ListOptions { get; set; }

        //题目编号
        public String TitleNumName { get; set; }

        //正确选项
        public String RightOptionName { get; set; }

        //专家指引
        public String ExpertsGuide { get; set; }

        public SingleChoiceTitle()
        {
            TitleName = "";
            TitleClassName = "";
            TitleNumName = "";
            RightOptionName = "";
            ExpertsGuide = "";
            if (ListOptions != null)
            {
                ListOptions.Clear();//清空数组
            }
            else
            {
                ListOptions = new List<string>();
            }
        }
        //是否是有效的题目
        public bool isVaild()
        {
            if (TitleName.Length > 0 && TitleClassName.Length > 0
                && TitleNumName.Length > 0 && ListOptions.Count() > 0
                && RightOptionName.Length > 0)
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
