namespace InstanGridMode
{
	partial class Export2ExcelHandmade
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
			if( disposing && (components != null) )
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Export2ExcelHandmade));
			this.bwExportExcel = new System.ComponentModel.BackgroundWorker();
			this.TimeElLabel = new DevExpress.XtraEditors.LabelControl();
			this.TimeLabel = new DevExpress.XtraEditors.LabelControl();
			this.btnCancel = new DevExpress.XtraEditors.SimpleButton();
			this.StatusLabel = new DevExpress.XtraEditors.LabelControl();
			this.progressBarControl = new DevExpress.XtraEditors.ProgressBarControl();
			((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// bwExportExcel
			// 
			this.bwExportExcel.WorkerReportsProgress = true;
			this.bwExportExcel.WorkerSupportsCancellation = true;
			this.bwExportExcel.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bwExportExcel_DoWork);
			this.bwExportExcel.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.bwExportExcel_ProgressChanged);
			this.bwExportExcel.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.bwExportExcel_RunWorkerCompleted);
			// 
			// TimeElLabel
			// 
			this.TimeElLabel.Location = new System.Drawing.Point(120, 45);
			this.TimeElLabel.Name = "TimeElLabel";
			this.TimeElLabel.Size = new System.Drawing.Size(28, 13);
			this.TimeElLabel.TabIndex = 18;
			this.TimeElLabel.Text = "00:00";
			// 
			// TimeLabel
			// 
			this.TimeLabel.Location = new System.Drawing.Point(12, 45);
			this.TimeLabel.Name = "TimeLabel";
			this.TimeLabel.Size = new System.Drawing.Size(87, 13);
			this.TimeLabel.TabIndex = 17;
			this.TimeLabel.Text = "Времени прошло:";
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(120, 93);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(75, 23);
			this.btnCancel.TabIndex = 19;
			this.btnCancel.Text = "Отмена";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// StatusLabel
			// 
			this.StatusLabel.AutoEllipsis = true;
			this.StatusLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
			this.StatusLabel.Location = new System.Drawing.Point(12, 12);
			this.StatusLabel.Name = "StatusLabel";
			this.StatusLabel.Size = new System.Drawing.Size(287, 17);
			this.StatusLabel.TabIndex = 20;
			this.StatusLabel.Text = "Выполняется экспорт базы в файл: Х";
			// 
			// progressBarControl
			// 
			this.progressBarControl.Location = new System.Drawing.Point(12, 69);
			this.progressBarControl.Name = "progressBarControl";
			this.progressBarControl.Size = new System.Drawing.Size(287, 18);
			this.progressBarControl.TabIndex = 21;
			// 
			// Export2ExcelHandmade
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(311, 128);
			this.Controls.Add(this.progressBarControl);
			this.Controls.Add(this.StatusLabel);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.TimeElLabel);
			this.Controls.Add(this.TimeLabel);
			this.MaximizeBox = false;
			this.MaximumSize = new System.Drawing.Size(327, 167);
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(327, 167);
			this.Name = "Export2ExcelHandmade";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Octarine v3.0";
			((System.ComponentModel.ISupportInitialize)(this.progressBarControl.Properties)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.ComponentModel.BackgroundWorker bwExportExcel;
		public DevExpress.XtraEditors.LabelControl TimeElLabel;
		private DevExpress.XtraEditors.LabelControl TimeLabel;
		private DevExpress.XtraEditors.SimpleButton btnCancel;
		private DevExpress.XtraEditors.LabelControl StatusLabel;
		private DevExpress.XtraEditors.ProgressBarControl progressBarControl;
	}
}
