using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace HoleLensDaemonProcess
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Process myProc = new Process();
            try
            {
                myProc.StartInfo.WorkingDirectory = System.Environment.CurrentDirectory;
                myProc.StartInfo.FileName = "HoloLensViewPlayer.exe";
                myProc.StartInfo.CreateNoWindow = true;  //不创建窗口
                myProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//这里设置DOS窗口不显示，经实践可行
                myProc.Start();
            }
            catch (Exception)
            {
                Console.WriteLine("HoloLensViewPlayer未找到");
            }

            {
                System.Threading.Thread.Sleep(2000);
                Console.WriteLine("休眠两秒检测进程是否存在");
            }
            while (myProc==null)
            {
                Console.WriteLine("操作系统进程不存在");
                System.Threading.Thread.Sleep(1000);
                try
                {
                    myProc.StartInfo.WorkingDirectory = System.Environment.CurrentDirectory;
                    myProc.StartInfo.FileName = "HoloLensViewPlayer.exe";
                    myProc.StartInfo.CreateNoWindow = true;  //不创建窗口
                    myProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//这里设置DOS窗口不显示，经实践可行
                    myProc.Start();
                    Console.WriteLine("重新启动应用程序");
                }
                catch (Exception)
                {
                    Console.WriteLine("HoloLensViewPlayer未找到");
                }
            }
            Console.WriteLine("应用程序已存在");
            System.IO.Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
            while (true)
            {
                System.Threading.Thread.Sleep(1000);
                Process[] ps = Process.GetProcessesByName("HoloLensSheelProcess");
                if (ps.Length > 0)
                {
                    Console.WriteLine("应用程序存在，这不做任何事情");
                }
                else
                {
                    myProc.CloseMainWindow();
                }
            }

        }
    }
}
