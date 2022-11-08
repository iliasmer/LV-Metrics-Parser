
namespace LV_Metrics_Parser
{
    partial class MainMenu
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainMenu));
            this.importSourceBtn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.sourceFileLabelPath = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // importSourceBtn
            // 
            this.importSourceBtn.BackColor = System.Drawing.Color.Teal;
            this.importSourceBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.importSourceBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.importSourceBtn.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.importSourceBtn.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.importSourceBtn.Location = new System.Drawing.Point(302, 186);
            this.importSourceBtn.Name = "importSourceBtn";
            this.importSourceBtn.Size = new System.Drawing.Size(159, 46);
            this.importSourceBtn.TabIndex = 3;
            this.importSourceBtn.Text = "ΧΡΗΣΗ";
            this.importSourceBtn.UseVisualStyleBackColor = false;
            this.importSourceBtn.Click += new System.EventHandler(this.importSourceBtn_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Teal;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Segoe UI Semibold", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.button1.ForeColor = System.Drawing.SystemColors.InactiveBorder;
            this.button1.Location = new System.Drawing.Point(302, 288);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(159, 46);
            this.button1.TabIndex = 4;
            this.button1.Text = "ΤΙΜΟΛΟΓΙΟ";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // sourceFileLabelPath
            // 
            this.sourceFileLabelPath.AutoSize = true;
            this.sourceFileLabelPath.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.sourceFileLabelPath.ForeColor = System.Drawing.SystemColors.WindowText;
            this.sourceFileLabelPath.Location = new System.Drawing.Point(230, 97);
            this.sourceFileLabelPath.Name = "sourceFileLabelPath";
            this.sourceFileLabelPath.Size = new System.Drawing.Size(320, 28);
            this.sourceFileLabelPath.TabIndex = 5;
            this.sourceFileLabelPath.Text = "Επεξεργασία Δεδομένων Ως Προς:";
            // 
            // MainMenu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.sourceFileLabelPath);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.importSourceBtn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MainMenu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Low Voltage Metrics Parser - Main Menu";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button importSourceBtn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label sourceFileLabelPath;
    }
}