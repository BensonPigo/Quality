using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoRemoveBackgroundProcess
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string run_type;

            if (args.Length > 0)
            {
                run_type = args[0];
            }
            else
            {
                run_type = "";
            }
            Sci.Env.AppInit();
            //若沒有傳入參數則正常啟動
            if (run_type.Equals(""))
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form(run_type));
            }
            else
            {
                //Application.Run(new Processing(run_type));
                Form installProcess = new Form(run_type);
                Thread.Sleep(5000);
            }
        }
    }
}
