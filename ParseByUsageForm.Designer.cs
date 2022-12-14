
namespace LV_Metrics_Parser
{
    partial class ParseByUsageForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ParseByUsageForm));
            this.monthSelectbox = new System.Windows.Forms.ComboBox();
            this.monthLabel = new System.Windows.Forms.Label();
            this.importSourceBtn = new System.Windows.Forms.Button();
            this.sourceFileLabelName = new System.Windows.Forms.Label();
            this.importSourceFile = new System.Windows.Forms.OpenFileDialog();
            this.sourceFileLabelPath = new System.Windows.Forms.Label();
            this.importFileProgressLabel = new System.Windows.Forms.Label();
            this.ExportBtnDEI = new System.Windows.Forms.Button();
            this.exportFilePathLabelDEI = new System.Windows.Forms.Label();
            this.exportFileNameLabelDEI = new System.Windows.Forms.Label();
            this.exportFileSourceDEI = new System.Windows.Forms.OpenFileDialog();
            this.exportFileProgressLabelDEI = new System.Windows.Forms.Label();
            this.exportFileProgressLabelPKY = new System.Windows.Forms.Label();
            this.exportFileNameLabelPKY = new System.Windows.Forms.Label();
            this.exportFilePathLabelPKY = new System.Windows.Forms.Label();
            this.ExportBtnPKY = new System.Windows.Forms.Button();
            this.exportFileSourcePKY = new System.Windows.Forms.OpenFileDialog();
            this.goDEIbtn = new System.Windows.Forms.Button();
            this.goPKYbtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // monthSelectbox
            // 
            this.monthSelectbox.BackColor = System.Drawing.SystemColors.HighlightText;
            this.monthSelectbox.Cursor = System.Windows.Forms.Cursors.Hand;
            this.monthSelectbox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.monthSelectbox.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.monthSelectbox.FormattingEnabled = true;
            this.monthSelectbox.Items.AddRange(new object[] {
            "-",
            "Ιανουάριος",
            "Φεβρουάριος",
            "Μάρτιος",
            "Απρίλιος",
            "Μάιος",
            "Ιούνιος",
            "Ιούλιος",
            "Αύγουστος",
            "Σεπτέμβριος",
            "Οκτώβριος",
            "Νοέμβριος",
            "Δεκέμβριος"});
            this.monthSelectbox.Location = new System.Drawing.Point(150, 250);
            this.monthSelectbox.Name = "monthSelectbox";
            this.monthSelectbox.Size = new System.Drawing.Size(276, 28);
            this.monthSelectbox.TabIndex = 0;
            this.monthSelectbox.SelectedIndexChanged += new System.EventHandler(this.monthSelectbox_SelectedIndexChanged);
            // 
            // monthLabel
            // 
            this.monthLabel.AutoSize = true;
            this.monthLabel.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.monthLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.monthLabel.Location = new System.Drawing.Point(144, 217);
            this.monthLabel.Name = "monthLabel";
            this.monthLabel.Size = new System.Drawing.Size(327, 28);
            this.monthLabel.TabIndex = 1;
            this.monthLabel.Text = "Επιλέξτε τον μήνα των μετρήσεων";
            // 
            // importSourceBtn
            // 
            this.importSourceBtn.BackColor = System.Drawing.Color.Teal;
            this.importSourceBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.importSourceBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.importSourceBtn.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.importSourceBtn.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.importSourceBtn.Location = new System.Drawing.Point(150, 116);
            this.importSourceBtn.Name = "importSourceBtn";
            this.importSourceBtn.Size = new System.Drawing.Size(159, 46);
            this.importSourceBtn.TabIndex = 2;
            this.importSourceBtn.Text = "Import";
            this.importSourceBtn.UseVisualStyleBackColor = false;
            this.importSourceBtn.Click += new System.EventHandler(this.importSourceBtn_Click);
            // 
            // sourceFileLabelName
            // 
            this.sourceFileLabelName.AutoSize = true;
            this.sourceFileLabelName.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.sourceFileLabelName.ForeColor = System.Drawing.SystemColors.WindowText;
            this.sourceFileLabelName.Location = new System.Drawing.Point(144, 36);
            this.sourceFileLabelName.Name = "sourceFileLabelName";
            this.sourceFileLabelName.Size = new System.Drawing.Size(99, 28);
            this.sourceFileLabelName.TabIndex = 3;
            this.sourceFileLabelName.Text = "File Name";
            // 
            // importSourceFile
            // 
            this.importSourceFile.FileName = "importSourceFile";
            // 
            // sourceFileLabelPath
            // 
            this.sourceFileLabelPath.AutoSize = true;
            this.sourceFileLabelPath.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.sourceFileLabelPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.sourceFileLabelPath.Location = new System.Drawing.Point(144, 76);
            this.sourceFileLabelPath.Name = "sourceFileLabelPath";
            this.sourceFileLabelPath.Size = new System.Drawing.Size(352, 28);
            this.sourceFileLabelPath.TabIndex = 4;
            this.sourceFileLabelPath.Text = "Επιλέξτε το αρχικό αρχείο με τις τιμές";
            // 
            // importFileProgressLabel
            // 
            this.importFileProgressLabel.AutoSize = true;
            this.importFileProgressLabel.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.importFileProgressLabel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.importFileProgressLabel.Location = new System.Drawing.Point(361, 125);
            this.importFileProgressLabel.Name = "importFileProgressLabel";
            this.importFileProgressLabel.Size = new System.Drawing.Size(39, 28);
            this.importFileProgressLabel.TabIndex = 5;
            this.importFileProgressLabel.Text = "0%";
            // 
            // ExportBtnDEI
            // 
            this.ExportBtnDEI.BackColor = System.Drawing.Color.Teal;
            this.ExportBtnDEI.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ExportBtnDEI.Enabled = false;
            this.ExportBtnDEI.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExportBtnDEI.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.ExportBtnDEI.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.ExportBtnDEI.Location = new System.Drawing.Point(150, 401);
            this.ExportBtnDEI.Name = "ExportBtnDEI";
            this.ExportBtnDEI.Size = new System.Drawing.Size(159, 46);
            this.ExportBtnDEI.TabIndex = 6;
            this.ExportBtnDEI.Text = "Export To";
            this.ExportBtnDEI.UseVisualStyleBackColor = false;
            this.ExportBtnDEI.Click += new System.EventHandler(this.ExportBtn_Click);
            // 
            // exportFilePathLabelDEI
            // 
            this.exportFilePathLabelDEI.AutoSize = true;
            this.exportFilePathLabelDEI.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFilePathLabelDEI.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFilePathLabelDEI.Location = new System.Drawing.Point(144, 358);
            this.exportFilePathLabelDEI.Name = "exportFilePathLabelDEI";
            this.exportFilePathLabelDEI.Size = new System.Drawing.Size(400, 28);
            this.exportFilePathLabelDEI.TabIndex = 7;
            this.exportFilePathLabelDEI.Text = "Επιλέξτε το αρχικό αρχείο για τις τιμές ΔΕΗ";
            // 
            // exportFileNameLabelDEI
            // 
            this.exportFileNameLabelDEI.AutoSize = true;
            this.exportFileNameLabelDEI.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFileNameLabelDEI.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFileNameLabelDEI.Location = new System.Drawing.Point(144, 318);
            this.exportFileNameLabelDEI.Name = "exportFileNameLabelDEI";
            this.exportFileNameLabelDEI.Size = new System.Drawing.Size(99, 28);
            this.exportFileNameLabelDEI.TabIndex = 8;
            this.exportFileNameLabelDEI.Text = "File Name";
            // 
            // exportFileProgressLabelDEI
            // 
            this.exportFileProgressLabelDEI.AutoSize = true;
            this.exportFileProgressLabelDEI.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFileProgressLabelDEI.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFileProgressLabelDEI.Location = new System.Drawing.Point(329, 410);
            this.exportFileProgressLabelDEI.Name = "exportFileProgressLabelDEI";
            this.exportFileProgressLabelDEI.Size = new System.Drawing.Size(39, 28);
            this.exportFileProgressLabelDEI.TabIndex = 9;
            this.exportFileProgressLabelDEI.Text = "0%";
            // 
            // exportFileProgressLabelPKY
            // 
            this.exportFileProgressLabelPKY.AutoSize = true;
            this.exportFileProgressLabelPKY.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFileProgressLabelPKY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFileProgressLabelPKY.Location = new System.Drawing.Point(329, 592);
            this.exportFileProgressLabelPKY.Name = "exportFileProgressLabelPKY";
            this.exportFileProgressLabelPKY.Size = new System.Drawing.Size(39, 28);
            this.exportFileProgressLabelPKY.TabIndex = 13;
            this.exportFileProgressLabelPKY.Text = "0%";
            // 
            // exportFileNameLabelPKY
            // 
            this.exportFileNameLabelPKY.AutoSize = true;
            this.exportFileNameLabelPKY.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFileNameLabelPKY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFileNameLabelPKY.Location = new System.Drawing.Point(150, 491);
            this.exportFileNameLabelPKY.Name = "exportFileNameLabelPKY";
            this.exportFileNameLabelPKY.Size = new System.Drawing.Size(99, 28);
            this.exportFileNameLabelPKY.TabIndex = 12;
            this.exportFileNameLabelPKY.Text = "File Name";
            // 
            // exportFilePathLabelPKY
            // 
            this.exportFilePathLabelPKY.AutoSize = true;
            this.exportFilePathLabelPKY.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.exportFilePathLabelPKY.ForeColor = System.Drawing.SystemColors.WindowText;
            this.exportFilePathLabelPKY.Location = new System.Drawing.Point(150, 531);
            this.exportFilePathLabelPKY.Name = "exportFilePathLabelPKY";
            this.exportFilePathLabelPKY.Size = new System.Drawing.Size(400, 28);
            this.exportFilePathLabelPKY.TabIndex = 11;
            this.exportFilePathLabelPKY.Text = "Επιλέξτε το αρχικό αρχείο για τις τιμές ΠΚΥ";
            // 
            // ExportBtnPKY
            // 
            this.ExportBtnPKY.BackColor = System.Drawing.Color.Teal;
            this.ExportBtnPKY.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ExportBtnPKY.Enabled = false;
            this.ExportBtnPKY.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExportBtnPKY.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.ExportBtnPKY.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.ExportBtnPKY.Location = new System.Drawing.Point(156, 574);
            this.ExportBtnPKY.Name = "ExportBtnPKY";
            this.ExportBtnPKY.Size = new System.Drawing.Size(153, 46);
            this.ExportBtnPKY.TabIndex = 10;
            this.ExportBtnPKY.Text = "Export To";
            this.ExportBtnPKY.UseVisualStyleBackColor = false;
            this.ExportBtnPKY.Click += new System.EventHandler(this.ExportBtnPKY_Click);
            // 
            // goDEIbtn
            // 
            this.goDEIbtn.BackColor = System.Drawing.Color.Teal;
            this.goDEIbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.goDEIbtn.Enabled = false;
            this.goDEIbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.goDEIbtn.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.goDEIbtn.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.goDEIbtn.Location = new System.Drawing.Point(569, 401);
            this.goDEIbtn.Name = "goDEIbtn";
            this.goDEIbtn.Size = new System.Drawing.Size(104, 46);
            this.goDEIbtn.TabIndex = 14;
            this.goDEIbtn.Text = "GO";
            this.goDEIbtn.UseVisualStyleBackColor = false;
            this.goDEIbtn.Click += new System.EventHandler(this.goDEIbtn_Click);
            // 
            // goPKYbtn
            // 
            this.goPKYbtn.BackColor = System.Drawing.Color.Teal;
            this.goPKYbtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.goPKYbtn.Enabled = false;
            this.goPKYbtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.goPKYbtn.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.goPKYbtn.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.goPKYbtn.Location = new System.Drawing.Point(569, 574);
            this.goPKYbtn.Name = "goPKYbtn";
            this.goPKYbtn.Size = new System.Drawing.Size(104, 46);
            this.goPKYbtn.TabIndex = 15;
            this.goPKYbtn.Text = "GO";
            this.goPKYbtn.UseVisualStyleBackColor = false;
            this.goPKYbtn.Click += new System.EventHandler(this.goPKYbtn_Click);
            // 
            // ParseByUsageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(828, 711);
            this.Controls.Add(this.goPKYbtn);
            this.Controls.Add(this.goDEIbtn);
            this.Controls.Add(this.exportFileProgressLabelPKY);
            this.Controls.Add(this.exportFileNameLabelPKY);
            this.Controls.Add(this.exportFilePathLabelPKY);
            this.Controls.Add(this.ExportBtnPKY);
            this.Controls.Add(this.exportFileProgressLabelDEI);
            this.Controls.Add(this.exportFileNameLabelDEI);
            this.Controls.Add(this.exportFilePathLabelDEI);
            this.Controls.Add(this.ExportBtnDEI);
            this.Controls.Add(this.importFileProgressLabel);
            this.Controls.Add(this.sourceFileLabelPath);
            this.Controls.Add(this.sourceFileLabelName);
            this.Controls.Add(this.importSourceBtn);
            this.Controls.Add(this.monthLabel);
            this.Controls.Add(this.monthSelectbox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ParseByUsageForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Low Voltage Metrics Parser - by Usage";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox monthSelectbox;
        private System.Windows.Forms.Label monthLabel;
        private System.Windows.Forms.Button importSourceBtn;
        private System.Windows.Forms.Label sourceFileLabelName;
        private System.Windows.Forms.OpenFileDialog importSourceFile;
        private System.Windows.Forms.Label sourceFileLabelPath;
        private System.Windows.Forms.Label importFileProgressLabel;
        private System.Windows.Forms.Button ExportBtnDEI;
        private System.Windows.Forms.Label exportFilePathLabelDEI;
        private System.Windows.Forms.Label exportFileNameLabelDEI;
        private System.Windows.Forms.OpenFileDialog exportFileSourceDEI;
        private System.Windows.Forms.Label exportFileProgressLabelDEI;
        private System.Windows.Forms.Label exportFileProgressLabelPKY;
        private System.Windows.Forms.Label exportFileNameLabelPKY;
        private System.Windows.Forms.Label exportFilePathLabelPKY;
        private System.Windows.Forms.Button ExportBtnPKY;
        private System.Windows.Forms.OpenFileDialog exportFileSourcePKY;
        private System.Windows.Forms.Button goDEIbtn;
        private System.Windows.Forms.Button goPKYbtn;
    }
}

