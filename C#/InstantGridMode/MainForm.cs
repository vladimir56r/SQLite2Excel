using DevExpress.Xpo;
using DevExpress.Xpo.Metadata;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace InstanGridMode
{
	public partial class MainForm : DevExpress.XtraEditors.XtraForm
	{
		string DBFileName = "test.db";
		int[] minColWidth;
		const int minWidth = 100;
		public MainForm()
		{
			InitializeComponent();
			loadDB();
		}
        void instantDS_ResolveSession(object sender, ResolveSessionEventArgs e)
		{
			Session session = new Session();
			session.ConnectionString = $@"XpoProvider=SQLite;Data Source={DBFileName};Read Only=True;";
			session.AutoCreateOption = DevExpress.Xpo.DB.AutoCreateOption.None;
			session.LockingOption = LockingOption.None;
			session.Connect();
			e.Session = session;
			try {
				using( var reader = new StreamReader("before_load.sql") )
				{
					var commandSQL = new System.Data.SQLite.SQLiteCommand();
					commandSQL.Connection = (System.Data.SQLite.SQLiteConnection)session.Connection;
					commandSQL.CommandText = reader.ReadToEnd();
					commandSQL.ExecuteNonQuery();
				}
			}catch(Exception ex){
				Console.WriteLine(ex.Message);
			}
		}   
		void instantDS_DismissSession(object sender, ResolveSessionEventArgs e)
		{
			IDisposable session = e.Session as IDisposable;
			if( session != null )
			{
				session.Dispose();
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			var fileName = Path.GetFileNameWithoutExtension(DBFileName);// +".xlsx";
			SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel (2007) (.xlsx) | *.xlsx|Все файлы (*.*)|*.*", AddExtension = true, DefaultExt = "xlsx", Title = "Сохранить как", FileName = fileName };
			if( sfd.ShowDialog() == DialogResult.OK )
			{
				Export2ExcelHandmade form = new Export2ExcelHandmade(gridView, DBFileName, "Test", sfd.FileName);
				form.ShowDialog();
				if( form.exportError )
					return;
				System.TimeSpan timeSpan = System.TimeSpan.FromSeconds((double)form.second);
				var timeAmount = string.Format("{0:D2}:{1:D2}", timeSpan.Minutes, timeSpan.Seconds);
				MessageBox.Show(this, "Экспортировано строк - " + gridView.RowCount + "\nВремени прошло - " + (new DateTime()).AddSeconds(form.second).ToLongTimeString(), "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
			}
		}

		private void loadDB()
		{
			string connectionString = $@"Data Source={DBFileName}";

			System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(connectionString);
			System.Data.SQLite.SQLiteCommand cmd = new System.Data.SQLite.SQLiteCommand("select * from test limit 1");
			cmd.Connection = conn;

			conn.Open();
			cmd.ExecuteScalar();
			System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(cmd);
			System.Data.DataSet ds = new System.Data.DataSet();

			da.Fill(ds);
			var table = ds.Tables[0];
			minColWidth = new int[table.Columns.Count];
			for( var colIndex = 0; colIndex < minColWidth.Length; colIndex++ )
				minColWidth[colIndex] = minWidth;
				foreach(DataRow row in table.Rows )
				for(var colIndex = 0; colIndex < minColWidth.Length; colIndex++ )
					minColWidth[colIndex] = Math.Max(minColWidth[colIndex], TextRenderer.MeasureText(row[colIndex].ToString(), gridControl1.Font).Width);
			ReflectionDictionary dict = new ReflectionDictionary();
			XpoDefault.Dictionary = dict;
			XPClassInfo classInfo = new XPDataObjectClassInfo(dict, "Test", new Attribute[] { new OptimisticLockingAttribute(false), new DeferredDeletionAttribute(false) });
			List<string> colNames = new List<string>();
			foreach( DataColumn col in table.Columns )
			{
				colNames.Add(col.ColumnName);
				classInfo.CreateMember(col.ColumnName, col.DataType);
			}
			classInfo.GetMember("id").AddAttribute(new KeyAttribute());

			XPInstantFeedbackSource instantDS = new XPInstantFeedbackSource(classInfo);
			instantDS.ResolveSession += instantDS_ResolveSession;
			instantDS.DismissSession += instantDS_DismissSession;
			gridView.Columns.Clear();
			gridControl1.DataSource = instantDS;
			gridControl1.Refresh();
			btnBestFirColumns.PerformClick();
		}

		private void OpenDB(object sender, EventArgs e)
		{
			var fileName = Path.GetFileNameWithoutExtension(DBFileName);
			OpenFileDialog sfd = new OpenFileDialog() { Filter = "SQLite database (.db) | *.db", AddExtension = true, DefaultExt = "db", Title = "Открыть базу данных", FileName = fileName };
			if( sfd.ShowDialog() == DialogResult.OK )
			{
				DBFileName = sfd.FileName;
				loadDB();
			}
		}

		private void BestFitColumns(object sender, EventArgs e)
		{
			var i = 0;
			foreach( DevExpress.XtraGrid.Columns.GridColumn column in gridView.Columns )
			{
				column.MinWidth =
				column.Width = minColWidth[i++];
			}
		}
	}
}