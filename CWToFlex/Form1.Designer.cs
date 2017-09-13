namespace CWToFlex
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dtpFrom = new System.Windows.Forms.DateTimePicker();
            this.dtpTom = new System.Windows.Forms.DateTimePicker();
            this.lbldtpFrom = new System.Windows.Forms.Label();
            this.lbldtpTom = new System.Windows.Forms.Label();
            this.btnGetCWData = new System.Windows.Forms.Button();
            this.lbCWData = new System.Windows.Forms.ListBox();
            this.txtSaveFileToDir = new System.Windows.Forms.TextBox();
            this.btnChangeSaveDir = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCreateFile = new System.Windows.Forms.Button();
            this.ss1 = new System.Windows.Forms.StatusStrip();
            this.SuspendLayout();
            // 
            // dtpFrom
            // 
            resources.ApplyResources(this.dtpFrom, "dtpFrom");
            this.dtpFrom.Name = "dtpFrom";
            // 
            // dtpTom
            // 
            resources.ApplyResources(this.dtpTom, "dtpTom");
            this.dtpTom.Name = "dtpTom";
            // 
            // lbldtpFrom
            // 
            resources.ApplyResources(this.lbldtpFrom, "lbldtpFrom");
            this.lbldtpFrom.Name = "lbldtpFrom";
            // 
            // lbldtpTom
            // 
            resources.ApplyResources(this.lbldtpTom, "lbldtpTom");
            this.lbldtpTom.Name = "lbldtpTom";
            // 
            // btnGetCWData
            // 
            resources.ApplyResources(this.btnGetCWData, "btnGetCWData");
            this.btnGetCWData.Name = "btnGetCWData";
            this.btnGetCWData.UseVisualStyleBackColor = true;
            this.btnGetCWData.Click += new System.EventHandler(this.btnGetCWData_Click);
            // 
            // lbCWData
            // 
            resources.ApplyResources(this.lbCWData, "lbCWData");
            this.lbCWData.FormattingEnabled = true;
            this.lbCWData.Name = "lbCWData";
            // 
            // txtSaveFileToDir
            // 
            resources.ApplyResources(this.txtSaveFileToDir, "txtSaveFileToDir");
            this.txtSaveFileToDir.Name = "txtSaveFileToDir";
            // 
            // btnChangeSaveDir
            // 
            resources.ApplyResources(this.btnChangeSaveDir, "btnChangeSaveDir");
            this.btnChangeSaveDir.Name = "btnChangeSaveDir";
            this.btnChangeSaveDir.UseVisualStyleBackColor = true;
            this.btnChangeSaveDir.Click += new System.EventHandler(this.btnChangeSaveDir_Click);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // btnCreateFile
            // 
            resources.ApplyResources(this.btnCreateFile, "btnCreateFile");
            this.btnCreateFile.Name = "btnCreateFile";
            this.btnCreateFile.UseVisualStyleBackColor = true;
            this.btnCreateFile.Click += new System.EventHandler(this.btnCreateFile_Click);
            // 
            // ss1
            // 
            resources.ApplyResources(this.ss1, "ss1");
            this.ss1.Name = "ss1";
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ss1);
            this.Controls.Add(this.btnCreateFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnChangeSaveDir);
            this.Controls.Add(this.txtSaveFileToDir);
            this.Controls.Add(this.lbCWData);
            this.Controls.Add(this.btnGetCWData);
            this.Controls.Add(this.lbldtpTom);
            this.Controls.Add(this.lbldtpFrom);
            this.Controls.Add(this.dtpTom);
            this.Controls.Add(this.dtpFrom);
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtpFrom;
        private System.Windows.Forms.DateTimePicker dtpTom;
        private System.Windows.Forms.Label lbldtpFrom;
        private System.Windows.Forms.Label lbldtpTom;
        private System.Windows.Forms.Button btnGetCWData;
        private System.Windows.Forms.ListBox lbCWData;
        private System.Windows.Forms.TextBox txtSaveFileToDir;
        private System.Windows.Forms.Button btnChangeSaveDir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCreateFile;
        private System.Windows.Forms.StatusStrip ss1;
    }
}

