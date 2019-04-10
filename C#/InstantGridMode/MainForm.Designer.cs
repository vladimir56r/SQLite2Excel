namespace InstanGridMode
{
	partial class MainForm
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
			this.panel1 = new System.Windows.Forms.Panel();
			this.btnBestFirColumns = new System.Windows.Forms.Button();
			this.btnOpenDb = new System.Windows.Forms.Button();
			this.btnExportToExcel = new System.Windows.Forms.Button();
			this.gridControl1 = new DevExpress.XtraGrid.GridControl();
			this.gridView = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.entityInstantFeedbackSource1 = new DevExpress.Data.Linq.EntityInstantFeedbackSource();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView)).BeginInit();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.btnBestFirColumns);
			this.panel1.Controls.Add(this.btnOpenDb);
			this.panel1.Controls.Add(this.btnExportToExcel);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(632, 57);
			this.panel1.TabIndex = 1;
			// 
			// btnBestFirColumns
			// 
			this.btnBestFirColumns.Location = new System.Drawing.Point(284, 11);
			this.btnBestFirColumns.Name = "btnBestFirColumns";
			this.btnBestFirColumns.Size = new System.Drawing.Size(130, 34);
			this.btnBestFirColumns.TabIndex = 2;
			this.btnBestFirColumns.Text = "Автоподбор ширины столбцов";
			this.btnBestFirColumns.UseVisualStyleBackColor = true;
			this.btnBestFirColumns.Click += new System.EventHandler(this.BestFitColumns);
			// 
			// btnOpenDb
			// 
			this.btnOpenDb.Location = new System.Drawing.Point(12, 11);
			this.btnOpenDb.Name = "btnOpenDb";
			this.btnOpenDb.Size = new System.Drawing.Size(130, 34);
			this.btnOpenDb.TabIndex = 0;
			this.btnOpenDb.Text = "Открыть базу";
			this.btnOpenDb.UseVisualStyleBackColor = true;
			this.btnOpenDb.Click += new System.EventHandler(this.OpenDB);
			// 
			// btnExportToExcel
			// 
			this.btnExportToExcel.Location = new System.Drawing.Point(148, 11);
			this.btnExportToExcel.Name = "btnExportToExcel";
			this.btnExportToExcel.Size = new System.Drawing.Size(130, 34);
			this.btnExportToExcel.TabIndex = 1;
			this.btnExportToExcel.Text = "Экспорт в Excel";
			this.btnExportToExcel.UseVisualStyleBackColor = true;
			this.btnExportToExcel.Click += new System.EventHandler(this.button1_Click);
			// 
			// gridControl1
			// 
			this.gridControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.gridControl1.Location = new System.Drawing.Point(0, 57);
			this.gridControl1.MainView = this.gridView;
			this.gridControl1.Name = "gridControl1";
			this.gridControl1.Size = new System.Drawing.Size(632, 454);
			this.gridControl1.TabIndex = 3;
			this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView});
			// 
			// gridView
			// 
			this.gridView.GridControl = this.gridControl1;
			this.gridView.Name = "gridView";
			this.gridView.OptionsDetail.DetailMode = DevExpress.XtraGrid.Views.Grid.DetailMode.Default;
			this.gridView.OptionsView.ColumnAutoWidth = false;
			this.gridView.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.True;
			this.gridView.OptionsView.ShowFooter = true;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(632, 511);
			this.Controls.Add(this.gridControl1);
			this.Controls.Add(this.panel1);
			this.Name = "MainForm";
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Button btnExportToExcel;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView;
		private DevExpress.XtraGrid.GridControl gridControl1;
		private DevExpress.Data.Linq.EntityInstantFeedbackSource entityInstantFeedbackSource1;
		private System.Windows.Forms.Button btnOpenDb;
		private System.Windows.Forms.Button btnBestFirColumns;
	}
}

