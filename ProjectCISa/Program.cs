using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectCIS
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
       // [DllImport("user32.dll", SetLastError = true)]
       // public static extern void SwitchToThisWindow(IntPtr hWnd, bool turnon);
        static void Main()
        {
            bool isReady;
            var mutex = new Mutex(true, "CISa", out isReady);
            if (!isReady)
            {
                //String frm = "CISa";
                // Process[] procs = Process.GetProcessesByName(frm);
                // foreach (Process proc in procs)
                //  {
                //     SwitchToThisWindow(proc.MainWindowHandle, true);
                //  }
                MessageBox.Show("ขณะนี้คุณได้เปิดโปรแกรมอยู่แล้ว!!","แจ้งเตือน",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                return;
            }
            GC.KeepAlive(mutex);
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new frmMain());
        }
    }
}
