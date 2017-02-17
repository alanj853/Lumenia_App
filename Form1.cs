using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
namespace WindowsFormsApplication1
{
    public partial class Application : Form
    {
        public Application()
        {
            InitializeComponent();
        }


        private String fileName = "No File Selected";
        private int startRow = 6;
        private int startColumn = 1;
        private int numberofScorers = 1;
        private Boolean appIsRunning = false;
        ConsoleApplication2.FunctionalReqScoreGen f;
        int exitCode = -1;

        private void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Hello World");
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Select an Excel File";


            // Show the Dialog.
            // If the user clicked OK in the dialog and
            // a .CUR file was selected, open it.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Assign the cursor in the Stream to the Form's Cursor property.
                fileName = openFileDialog1.FileName;
                // this.Cursor = new Cursor(openFileDialog1.OpenFile());
                //MessageBox.Show("The file you selected is " + fileName);
                //this.fileNameDisplay.Text = fileName;

                if (fileName != "No File Selected")
                {
                    this.fileNameDisplay.Text = fileName;
                    this.status.Text = "Waiting for User to Click 'Go'";
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void printBuildSummary()
        {
            int exitCode = -1;
            while (!f.appFinishedRunning)
                exitCode = f.getExitCode();

            exitCode = f.getExitCode();

            if (exitCode != 0)
            {
                status.Text = "An Error has Occured - See CommandLine Window for Details. (Exit Code = " + exitCode + ")";

            }
            else
                status.Text = "Finished Building Excel Table. Exit Code = " + exitCode;
            exitLabel.Text = "Process Completed";

            appIsRunning = false;
        }

        private void goButton_Click(object sender, EventArgs e)
        {
            if (!appIsRunning)
            {

                appIsRunning = true;
                //String msg = "File: " + fileName + "\n" + "Start Row = " + startRow + "\nStart Column = " + startColumn + "\nNumber Of Scorers = " + numberofScorers + "\n\nAre all of these parameters correct?\nIf yes, click 'Yes' to build table. If not, click 'No' to change them.";
                String msg = "File: " + fileName + "\n" + "\nNumber Of Scorers = " + numberofScorers + "\n\nAre all of these parameters correct?\nIf yes, click 'Yes' to build table. If not, click 'No' to change them.";
                /*DialogResult dialogResult = MessageBox.Show(msg, "Confirmation", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    status.Text = "Building Excel Table...";
                    exitLabel.Text = "Please do not close application";
                    f = new ConsoleApplication2.FunctionalReqScoreGen(fileName, numberofScorers, startRow, startColumn);
                    backgroundWorker1.RunWorkerAsync();
                    Thread t1 = new Thread(f.run);
                    t1.Start();
                    goButton.Enabled = false;

                }*/

                status.Text = "Building Excel Table...";
                exitLabel.Text = "Please do not close application";
                f = new ConsoleApplication2.FunctionalReqScoreGen(fileName, numberofScorers, startRow, startColumn);
                backgroundWorker1.RunWorkerAsync();
                Thread t1 = new Thread(f.run);
                t1.Start();
                goButton.Enabled = false;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            // startColumn = (int)numericUpDown1.Value;

        }

        private void numericUpDown1_Row_ValueChanged(object sender, EventArgs e)
        {

            //startRow = (int)numericUpDown1_Row.Value;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("File is valid");
        }

        private void numericUpDown_Scorers_ValueChanged(object sender, EventArgs e)
        {

            //this.noScorers = Convert.ToInt32(noScorers.Value);
            this.numberofScorers = (int)numericUpDown_Scorers.Value;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Application_Load(object sender, EventArgs e)
        {

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            double prog = 0.0;
            int i = 0;
            while (!f.appFinishedRunning)
            {
                prog = 100 * f.getCompletionProgress();
                i = (int)prog;
                backgroundWorker1.ReportProgress(i);
                Thread.Sleep(100);

            }
            prog = 100 * f.getCompletionProgress();
            backgroundWorker1.ReportProgress((int)prog);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            printBuildSummary();

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            appProgressBar.Value = e.ProgressPercentage;
            progLabel.Text = e.ProgressPercentage.ToString() + " %";
        }









    }

}
