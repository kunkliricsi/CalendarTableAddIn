namespace CalendarTableAddIn
{
    partial class DialogLauncher
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
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.monthPicker1 = new CalendarTableAddIn.MonthPicker();
            this.SuspendLayout();
            // 
            // monthPicker1
            // 
            this.monthPicker1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.monthPicker1.Location = new System.Drawing.Point(0, 0);
            this.monthPicker1.Name = "monthPicker1";
            this.monthPicker1.ShowToday = false;
            this.monthPicker1.ShowTodayCircle = false;
            this.monthPicker1.TabIndex = 0;
            // 
            // DialogLauncher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(230, 161);
            this.Controls.Add(this.monthPicker1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "DialogLauncher";
            this.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds;
            this.Text = "Pick a month";
            this.ResumeLayout(false);

        }

        #endregion
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private MonthPicker monthPicker1;
    }
}