using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lumenia_App
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try {
                Application.Run(new Form1());
                //throw new Exception("Dummy Exception");
            }
            catch (Exception e) {
                Console.WriteLine("Exception Caught");
                DateTime today = DateTime.Now;
                String time = today.Day.ToString() + today.Month.ToString() + today.Year.ToString() + "_" + today.Hour.ToString() + today.Minute.ToString() + today.Second.ToString();
                String filename = "error_log_program" + time + ".txt";
                String fileContent = e.Message.ToString() + Environment.NewLine;
                System.IO.File.WriteAllText(filename, fileContent);
            }
        }
    } 
}
        
