namespace Lumenia_App
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            //this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            //this.Text = "Form1";







            this.button1 = new System.Windows.Forms.Button();
            this.fileNameDisplay = new System.Windows.Forms.TextBox();
            this.goButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown_Scorers = new System.Windows.Forms.NumericUpDown();
            this.status = new System.Windows.Forms.TextBox();
            this.statusLabel = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.exitLabel = new System.Windows.Forms.Label();
            this.testLabel = new System.Windows.Forms.Label();
            this.appProgressBar = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Scorers)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(391, 19);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(121, 36);
            this.button1.TabIndex = 0;
            this.button1.Text = "Choose Excel File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // fileNameDisplay
            // 
            this.fileNameDisplay.Location = new System.Drawing.Point(22, 28);
            this.fileNameDisplay.Name = "fileNameDisplay";
            this.fileNameDisplay.Size = new System.Drawing.Size(340, 20);
            this.fileNameDisplay.TabIndex = 2;
            // 
            // goButton
            // 
            this.goButton.Location = new System.Drawing.Point(391, 189);
            this.goButton.Name = "goButton";
            this.goButton.Size = new System.Drawing.Size(121, 28);
            this.goButton.TabIndex = 5;
            this.goButton.Text = "Go";
            this.goButton.UseVisualStyleBackColor = true;
            this.goButton.Click += new System.EventHandler(this.goButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(388, 108);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Number of Scorers";
            // 
            // numericUpDown_Scorers
            // 
            this.numericUpDown_Scorers.Location = new System.Drawing.Point(391, 133);
            this.numericUpDown_Scorers.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown_Scorers.Name = "numericUpDown_Scorers";
            this.numericUpDown_Scorers.Size = new System.Drawing.Size(54, 20);
            this.numericUpDown_Scorers.TabIndex = 9;
            this.numericUpDown_Scorers.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown_Scorers.ValueChanged += new System.EventHandler(this.numericUpDown_Scorers_ValueChanged);
            // 
            // status
            // 
            this.status.Location = new System.Drawing.Point(19, 110);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(215, 20);
            this.status.TabIndex = 10;
            this.status.Text = "Waiting for User to select a File";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.BackColor = System.Drawing.SystemColors.HotTrack;
            this.statusLabel.Location = new System.Drawing.Point(19, 91);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(92, 13);
            this.statusLabel.TabIndex = 11;
            this.statusLabel.Text = "Application Status";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(391, 224);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(121, 28);
            this.button2.TabIndex = 12;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // exitLabel
            // 
            this.exitLabel.AllowDrop = true;
            this.exitLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exitLabel.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.exitLabel.Location = new System.Drawing.Point(15, 189);
            this.exitLabel.Name = "exitLabel";
            this.exitLabel.Size = new System.Drawing.Size(347, 63);
            this.exitLabel.TabIndex = 13;
            // 
            // testLabel
            // 
            this.testLabel.AutoSize = true;
            this.testLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.testLabel.Location = new System.Drawing.Point(12, 140);
            this.testLabel.Name = "testLabel";
            this.testLabel.Size = new System.Drawing.Size(0, 20);
            this.testLabel.TabIndex = 14;
            // 
            // appProgressBar
            // 
            this.appProgressBar.Location = new System.Drawing.Point(22, 153);
            this.appProgressBar.Name = "appProgressBar";
            this.appProgressBar.Size = new System.Drawing.Size(212, 23);
            this.appProgressBar.TabIndex = 15;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // progLabel
            // 
            this.progLabel.Location = new System.Drawing.Point(250, 153);
            this.progLabel.Name = "progLabel";
            this.progLabel.Size = new System.Drawing.Size(55, 23);
            this.progLabel.TabIndex = 16;
            this.progLabel.Text = "0 %";
            this.progLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Application
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(603, 261);
            this.Controls.Add(this.progLabel);
            this.Controls.Add(this.appProgressBar);
            this.Controls.Add(this.testLabel);
            this.Controls.Add(this.exitLabel);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.status);
            this.Controls.Add(this.numericUpDown_Scorers);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.goButton);
            this.Controls.Add(this.fileNameDisplay);
            this.Controls.Add(this.button1);
            this.Name = "Application";
            this.Text = "Functional Requirement Score Generator";
            this.Load += new System.EventHandler(this.Application_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Scorers)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.IO.FileSystemWatcher fileSystemWatcher1;
        private System.Windows.Forms.TextBox fileNameDisplay;
        private System.Windows.Forms.Button goButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown_Scorers;
        private System.Windows.Forms.TextBox status;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label exitLabel;
        private System.Windows.Forms.Label testLabel;
        private System.Windows.Forms.ProgressBar appProgressBar;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label progLabel;
    
    }
}

