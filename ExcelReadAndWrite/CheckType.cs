using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelReadAndWrite
{
    struct OPeratorState
    {
        public int Num;
        public String Value;
  
        public OPeratorState(int Num, String value)
        {
            this.Num = Num;
            this.Value = value;
        }
    };
    class CheckType
    {
        public Int32 Num { get; set; }
        public String LableName { get; set; }
        public String DataPortName { get; set; }
        public List<OPeratorState> OperatorStateList { get; set; }
        public bool UsePressLight { get; set; }
        public CheckType()
        {
            Num= -1;
            LableName="";
            DataPortName="";
            UsePressLight = false;
            if (OperatorStateList != null)
            {
                OperatorStateList.Clear();
            }
            else
            {
                OperatorStateList = new List<OPeratorState>();
                OperatorStateList.Clear();
            }
        }
        public bool IsValid()
        {
            if (LableName.Length > 0 && DataPortName.Length > 0 && Num != -1)
            {
                return true;
            }
            return false;
        }
        public bool WriteOperatorStateToList(String stateValue)
        {
            try {
                string[] TstateArray = stateValue.Split(new Char[]{ '\n' },StringSplitOptions.RemoveEmptyEntries);
                if (TstateArray.Count() > 1)
                {
                    OPeratorState _state;
                    foreach (string item in TstateArray)
                    {
                        string[] ItemOPeratorStateArray = item.Split(new Char[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
                        if (ItemOPeratorStateArray.Count() == 2)
                        {
                            _state.Num = int.Parse(ItemOPeratorStateArray[0]);
                            _state.Value = ItemOPeratorStateArray[1];
                            OperatorStateList.Add(_state);
                        }
                    }
                    return true;
                }
             
            }
            catch
            {
                Console.WriteLine("WriteOperatorStateToList=====ErrorName={0}", stateValue);
            }
            return false;
        }
    }
}
