using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace HoloLensSheelProcess
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Process myProc = new Process();
            try
            {
                myProc.StartInfo.WorkingDirectory = System.Environment.CurrentDirectory;
                myProc.StartInfo.FileName = "HoleLensDaemonProcess.exe";
                myProc.StartInfo.CreateNoWindow = true;  //不创建窗口
                myProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//这里设置DOS窗口不显示，经实践可行
                myProc.Start();
            }
            catch (Exception)
            {
                Console.WriteLine("HoleLensDaemonProcess进程找不到");
            }
            System.Threading.Thread.Sleep(2000);
            Console.WriteLine("休眠两秒检测进程是否存在");
            System.IO.Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);
            while (myProc == null)
            {
                Console.WriteLine("操作系统进程不存在");
                System.Threading.Thread.Sleep(3000);
                try
                {
                    myProc.StartInfo.WorkingDirectory = System.Environment.CurrentDirectory;
                    myProc.StartInfo.FileName = "HoleLensDaemonProcess.exe";
                    myProc.StartInfo.CreateNoWindow = true;  //不创建窗口
                    myProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;//这里设置DOS窗口不显示，经实践可行
                    myProc.Start();
                    Console.WriteLine("重新启动应用程序");
                }
                catch (Exception)
                {
                    Console.WriteLine("HoleLensDaemonProcess进程找不到");

                }
            }
            while (true)
            {
                System.Threading.Thread.Sleep(3000);
                Console.WriteLine("程序休眠不处理任何事情");
            }
        }
    }
}
