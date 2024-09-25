using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
namespace ConvertFileContent
{
    class Program
    {
        static Dictionary<string, string> DataDictionary = new Dictionary<string, string>();
        static void Main(string[] args)
        {
            String TxtFilePath = @"K:\Nodejs\Demo\ContentJson\new.txt";
            ReadFile(TxtFilePath);
            String JsoFilePath = @"K:\Nodejs\Demo\ContentJson\Contents.json";
            String JsoWriteFilePath = @"K:\Nodejs\Demo\ContentJson\Contents1.json";
            ReadWriteJson(JsoFilePath, JsoWriteFilePath);

            JsoFilePath = @"K:\Nodejs\Demo\ContentJson\Score.json";
            JsoWriteFilePath = @"K:\Nodejs\Demo\ContentJson\Score1.json";
            ReadWriteScoreJson(JsoFilePath, JsoWriteFilePath);
            Console.ReadKey();
        }
        static void ReadFile(String Path)
        {
            FileStream fileStream = new FileStream(Path, FileMode.Open, FileAccess.Read);
            byte[] bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, bytes.Length);
            string FileContent = Encoding.UTF8.GetString(bytes);

            //Console.WriteLine(FileContent);
            String[] FileArray = FileContent.Split(new char[2] { '\r','\n'}, StringSplitOptions.RemoveEmptyEntries);
            for(int i = 0;i< FileArray.Length;++i)
            {
                string CurrentFile = FileArray[i];
                String[] LineData= CurrentFile.Split(new char[4] { '#', '#', '#', '#' }, StringSplitOptions.RemoveEmptyEntries);
                if (LineData.Length == 1)
                {
                    String[] KeyData = CurrentFile.Split('+');
                    if (KeyData.Length == 2)
                    {
                        if (!DataDictionary.ContainsKey(KeyData[0]))
                        {
                            DataDictionary.Add(KeyData[0], CurrentFile);
                        }
                    }
                }
                else
                {
                    for (int j = 0; j < LineData.Length; ++j)
                    {
                        string CurrentData= LineData[j];
                        String[] KeyData = CurrentData.Split('+');
                        if (KeyData.Length == 2)
                        {
                            if (!DataDictionary.ContainsKey(KeyData[0]))
                            {
                                DataDictionary.Add(KeyData[0], CurrentData);
                            }
                        }
                    }
                }
            }

            fileStream.Close();
        }
        static void ReadWriteJson(String Path, String WritePath)
        {
            try
            {
                string path = Path;
                StreamReader streamReader = new StreamReader(path);
                string jsonStr = streamReader.ReadToEnd();
                dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(jsonStr);
                JToken token = jsonObj["itemDetails"];
                int TotalChangedNum = 0;
                foreach (JObject e in token)
                {
                    //Console.WriteLine("Step Index = {0}",e["index"]);
                    if(e["userCustom"] != null)
                    {
                        JToken token1 = e["userCustom"];
                        foreach (JObject e1 in token1)
                        {
                            string key = e1["key"].ToString();
                            if (key == "MuShi")
                            {
                                string value = e1["value"].ToString();
                                String[] ValueArray = value.Split('+');
                                if(ValueArray.Length > 1)
                                {
                                    
                                    if(DataDictionary.ContainsKey(ValueArray[0]))
                                    {
                                        string dicValue = "";
                                        DataDictionary.TryGetValue(ValueArray[0],out dicValue);
                                        dicValue += "+MS";
                                        if (value != dicValue)
                                        {
                                            TotalChangedNum += 1;
                                            Console.WriteLine("Change Step Index = {0},Key={1},OldValue={2},NewValue={3}", e["index"], ValueArray[0], value, dicValue);
                                        }
                                        e1["value"] = dicValue;
                                    }
                                }
                            }
                        }
                    }


                }

                Console.WriteLine("总共更新了 {0}条数据", TotalChangedNum);
                streamReader.Close();
                string output = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObj, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(WritePath, output);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
        static void ReadWriteScoreJson(String Path, String WritePath)
        {
            try
            {
                string path = Path;
                StreamReader streamReader = new StreamReader(path);
                string jsonStr = streamReader.ReadToEnd();
                dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(jsonStr);
                JToken token = jsonObj["scoreDetail"];
                int TotalChangedNum = 0;
                foreach (JObject e in token)
                {
                    //Console.WriteLine("Step Index = {0}",e["index"]);
                    if (e["userCustom"] != null)
                    {
                        JToken token1 = e["userCustom"];
                        foreach (JObject e1 in token1)
                        {
                            string key = e1["key"].ToString();
                            if (key == "MuShi")
                            {
                                string value = e1["value"].ToString();
                                String[] ValueArray = value.Split('+');
                                if (ValueArray.Length > 1)
                                {

                                    if (DataDictionary.ContainsKey(ValueArray[0]))
                                    {
                                        string dicValue = "";
                                        DataDictionary.TryGetValue(ValueArray[0], out dicValue);
                                        dicValue += "+MS";
                                        if (value != dicValue)
                                        {
                                            TotalChangedNum += 1;
                                            Console.WriteLine("Change Step Index = {0},Key={1},OldValue={2},NewValue={3}", e["index"], ValueArray[0], value, dicValue);
                                        }
                                        e1["value"] = dicValue;
                                    }
                                }
                            }
                        }
                    }


                }

                Console.WriteLine("总共更新了 {0}条数据", TotalChangedNum);
                streamReader.Close();
                string output = Newtonsoft.Json.JsonConvert.SerializeObject(jsonObj, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(WritePath, output);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
