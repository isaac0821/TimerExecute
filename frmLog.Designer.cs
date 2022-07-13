
namespace TimerExecute
{
    partial class frmLog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLog));
            this.lsbLog = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // lsbLog
            // 
            this.lsbLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lsbLog.FormattingEnabled = true;
            this.lsbLog.ItemHeight = 15;
            this.lsbLog.Location = new System.Drawing.Point(0, 0);
            this.lsbLog.Name = "lsbLog";
            this.lsbLog.Size = new System.Drawing.Size(359, 450);
            this.lsbLog.TabIndex = 0;
            // 
            // frmLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(359, 450);
            this.Controls.Add(this.lsbLog);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmLog";
            this.Text = "Timer Execute - v0.0.1";
            this.Load += new System.EventHandler(this.frmLog_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lsbLog;
    }
}