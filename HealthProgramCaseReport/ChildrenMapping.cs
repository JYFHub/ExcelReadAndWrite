using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HealthProgramCaseReport
{
    class ChildrenMapping
    {
        //选项名称
        public string ChilderenOPtionName { get; set; }

        //选项列表
        public List<string> ChildrenOPtionList { get; set; }

        public ChildrenMapping()
        {
            ChilderenOPtionName = "";
            if (ChildrenOPtionList != null)
            {
                ChildrenOPtionList.Clear();//清空数组
            }
            else
            {
                ChildrenOPtionList = new List<string>();
            }
        }
        //是否是有效的题目
        public bool isVaild()
        {
            if (ChilderenOPtionName.Length > 0 && ChildrenOPtionList.Count() > 0)
            {
                return true;
            }
            return false;
        }
    }
}
