using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Compression;
using System.Globalization;
using DevExpress.Data.Filtering;

namespace InstanGridMode
{
	public partial class Export2ExcelHandmade : DevExpress.XtraEditors.XtraForm
	{
		public int second;
		public bool exportError = true;
		GridView _gridView;
		string _connectionString;
		string _tableName;
		public Timer timer { get; set; }

		const int _maxRowsInFile = 500000;
		System.Threading.ManualResetEvent _pauseManager = new System.Threading.ManualResetEvent(true);

		public Export2ExcelHandmade(GridView gridView, string DBFileName, string tableName, string outputExcelFileName)
		{
			_gridView = gridView;
			_tableName = tableName;
			_connectionString = $@"Data Source={DBFileName}";
			if( _gridView == null || DBFileName == null || _gridView.RowCount == 0 )
			{
				MessageBox.Show(this, "Невозможно выгрузить пустые данные в excel файл!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			InitializeComponent();
			Text = "";
			InitializeTimer();
			StatusLabel.Text = $"Выполняется экспорт базы в файл: {Path.GetFileNameWithoutExtension(outputExcelFileName) + ".xlsx"}\n";
			bwExportExcel.RunWorkerAsync(outputExcelFileName);
		}

		private void bwExportExcel_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			timer.Stop();
			System.TimeSpan timeSpan = System.TimeSpan.FromSeconds((double)this.second);
			var timeAmount = string.Format("{0:D2}:{1:D2}", timeSpan.Minutes, timeSpan.Seconds);
			if( e.Result is Exception )
			{
				MessageBox.Show(this, $"Произошла ошибка во время выгрузки в excel файл\nВозможная причина:\n{(e.Result as Exception).Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			else
				exportError = !(bool)e.Result;
			Close();
		}

		public void InitializeTimer()
		{
			this.timer = new Timer() { Interval = 1000, Enabled = true };
			this.timer.Tick += new System.EventHandler((sender, e) => {
				System.TimeSpan timeSpan = System.TimeSpan.FromSeconds((double)this.second);
				second++;
				TimeElLabel.Text = string.Format("{0:D2}:{1:D2}", timeSpan.Minutes, timeSpan.Seconds);
			});
			timer.Start();
		}

		private void btnCancel_Click(object sender, EventArgs e)
		{
			_pauseManager.Reset();
			timer.Stop();
			var dialogResult = MessageBox.Show(this, $"Прервать экспорт?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			timer.Start();
			_pauseManager.Set();
			if( dialogResult == DialogResult.Yes )
			{
				StatusLabel.Text = $"Выполняется отмена экспорта...";
				btnCancel.Enabled = false;
				bwExportExcel.CancelAsync();
			}
		}

		static string NumberToLetters(int number)
		{
			string result;
			if( number > 0 )
			{
				int alphabets = (number - 1) / 26;
				int remainder = (number - 1) % 26;
				result = ((char)('A' + remainder)).ToString();
				if( alphabets > 0 )
					result = NumberToLetters(alphabets) + result;
			}
			else
				result = null;
			return result;
		}

		private void bwExportExcel_DoWork(object sender, DoWorkEventArgs e)
		{
			e.Result = true;
			var fname = e.Argument as string;
			var fnameWithoutExt = Path.GetFileNameWithoutExtension(fname);
			var fileExt = Path.GetExtension(fname);
			var outputDir = Path.GetDirectoryName(fname);
			try
			{
				if( File.Exists(fname) )
				{
					File.Delete(fname);
				}
				using( FileStream zipToOpen = new FileStream(fname, FileMode.OpenOrCreate) )
				{
					using( ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Create) )
					{
						// Dirs
						var xlDir = $@"xl";
						var worksheetDir = $@"{xlDir}\worksheets";
						var relsDir = $@"{xlDir}\_rels";
						var generalRelsDir = $@"_rels";
						// FNames
						var contentTypesFileName = $@"[Content_Types].xml";
						var workbookFileName = $@"{xlDir}\workbook.xml";
						var sharedStringFileName = $@"{xlDir}\sharedStrings.xml";
						var stylesFileName = $@"{xlDir}\styles.xml";
						var relsFileName = $@"{relsDir}\workbook.xml.rels";
						var generalRelsFileName = $@"{generalRelsDir}\.rels";

						// Create workbook, styles and other files
						ZipArchiveEntry readmeEntry = archive.CreateEntry(stylesFileName, CompressionLevel.Optimal);
						using( var stylesFile = new StreamWriter(readmeEntry.Open()) )
						{
							stylesFile.Write(excelTemplateStyles, excelTemplateStylesDateFormat);
						}
						readmeEntry = archive.CreateEntry(generalRelsFileName, CompressionLevel.Optimal);
						using( var generalRelsFile = new StreamWriter(readmeEntry.Open()) )
						{
							generalRelsFile.Write(excelTemplateGeneralRels);
						}
						readmeEntry = archive.CreateEntry(sharedStringFileName, CompressionLevel.Optimal);
						using( var sharedStringFile = new StreamWriter(readmeEntry.Open()) )
						{
							sharedStringFile.Write(excelTemplateSharedString);
						}

						var wbSheetsTemplBuilder = new StringBuilder();
						var wbSheetsTemplSheetDefNameBuilder = new StringBuilder();
						var ctWorkSheetOverrideBuilder = new StringBuilder();
						var relTemplBuilder = new StringBuilder();

						// Create worksheets
						var worksheetsCount = _gridView.RowCount / _maxRowsInFile + 1;
						using( SQLiteConnection conn = new SQLiteConnection(_connectionString) )
						{
							conn.Open();
							var SQLWhereFromGrid = CriteriaToWhereClauseHelper.GetOracleWhere(_gridView.ActiveFilterCriteria);
							var SQLColumns = string.Join(",", _gridView.VisibleColumns.Select(x => x.Name.Substring(3)));
							if( SQLColumns != "" )
							{
								var txtQuery = $"SELECT {SQLColumns} FROM {_tableName} {(SQLWhereFromGrid != "" ? ($"WHERE {SQLWhereFromGrid}") : "") }";
								using( SQLiteCommand cmd = new SQLiteCommand(txtQuery, conn) )
								using( SQLiteDataReader rd = cmd.ExecuteReader() )
									for( var worksheetIndex = 0; worksheetIndex < worksheetsCount; worksheetIndex++ )
									{
										var worksheetFileName = $@"{worksheetDir}\sheet{worksheetIndex + 1}.xml";
										var wsName = worksheetsCount > 1
													? $"Billing #{worksheetIndex + 1}"
													: "Billing";
										readmeEntry = archive.CreateEntry(worksheetFileName, CompressionLevel.Optimal);
										using( var worksheetFile = new StreamWriter(readmeEntry.Open()) )
										{
											var startRowInGrid = worksheetIndex * _maxRowsInFile;
											var rowsInCurrentFile =
												_gridView.RowCount - startRowInGrid > _maxRowsInFile
												? _maxRowsInFile
												: _gridView.RowCount - startRowInGrid;
											var endColumnsRange = NumberToLetters(_gridView.VisibleColumns.Count);

											// Add in builder cur sheet
											wbSheetsTemplBuilder.AppendFormat(excelTemplateWorkbookSheet, wsName, worksheetIndex + 1, worksheetIndex + 3);
											wbSheetsTemplSheetDefNameBuilder.AppendFormat(excelTemplateWorkbookSheetDefName, worksheetIndex, wsName, endColumnsRange, rowsInCurrentFile + 1);
											ctWorkSheetOverrideBuilder.AppendFormat(excelTemplateContentTypesWorkSheetOverride, worksheetIndex + 1);
											relTemplBuilder.AppendFormat(excelTemplateRelsForWorksheet, worksheetIndex + 3, worksheetIndex + 1);
											// Create header, coldata and open body in worksheet
											worksheetFile.Write(excelTemplateWorksheetHeader, endColumnsRange, rowsInCurrentFile + 1, 0);
											worksheetFile.Write(excelTemplateWorksheetColumn, _gridView.VisibleColumns.Count);
											worksheetFile.Write(excelTemplateWorksheetBeginBody);

											// Write header
											var j = 0;
											var rowBuilder = new StringBuilder();
											foreach( string colName in _gridView.VisibleColumns.Select(x => x.Name) )
											{
												j++;
												rowBuilder.Append(string.Format(excelTemplateWorksheetCol, NumberToLetters(j), 1, 2, "t =\"inlineStr\"", $"<is><t>{System.Net.WebUtility.HtmlEncode(colName.Substring(3))}</t></is>"));
											}
											worksheetFile.Write(excelTemplateWorksheetRow, 1, rowBuilder.ToString());

											// Export data rows						
											var curProgress = 0;

											for( var i = 0; i < rowsInCurrentFile; i++ )
											{
												var res = rd.Read();
												if( !res )
													break;
												rowBuilder = new StringBuilder();
												for( j = 0; j < rd.FieldCount; j++ )
												{
													var value = rd.GetValue(j);
													var valueType = rd.GetFieldType(j);
													// <c r ="{0}{1}" s="{2}" {3}>{4}</c>
													var isDate = (valueType == typeof(DateTime));
													var isInt = (valueType == typeof(int));
													var useInlineStr = !isDate && !isInt;
													rowBuilder.Append(string.Format(excelTemplateWorksheetCol,
														NumberToLetters(j + 1), i + 2,
														isDate
															? 3
															: isInt
																? 4
																: 1,
														useInlineStr
															? "t =\"inlineStr\""
															: "",
														useInlineStr
															? $"<is><t>{(value == null || value.GetType() == typeof(System.DBNull) ? "" : System.Net.WebUtility.HtmlEncode(value as string))}</t></is>"
															: $"<v>{ (value == null || value.GetType() == typeof(System.DBNull) ? "" : (isDate ? ((DateTime)value).ToOADate() : value))}</v>"
														));
												}
												worksheetFile.Write(excelTemplateWorksheetRow, i + 2, rowBuilder.ToString());

												if( bwExportExcel.CancellationPending )
												{
													e.Result = false;
													return;
												}

												_pauseManager.WaitOne();

												if( 100 * (startRowInGrid + i) / _gridView.RowCount > curProgress )
												{
													curProgress = 100 * (startRowInGrid + i) / _gridView.RowCount;
													bwExportExcel.ReportProgress(curProgress);
												}
											}
											worksheetFile.Write(excelTemplateWorksheetEndBody, endColumnsRange, rowsInCurrentFile + 1);
										}
									}
							}
						}
						readmeEntry = archive.CreateEntry(contentTypesFileName, CompressionLevel.Optimal);
						using( var contentTypesFile = new StreamWriter(readmeEntry.Open()) )
						{
							contentTypesFile.Write(excelTemplateContentTypes, ctWorkSheetOverrideBuilder.ToString());
						}
						readmeEntry = archive.CreateEntry(workbookFileName, CompressionLevel.Optimal);
						using( var workbookFile = new StreamWriter(readmeEntry.Open()) )
						{
							workbookFile.Write(excelTemplateWorkbook, wbSheetsTemplBuilder.ToString(), wbSheetsTemplSheetDefNameBuilder.ToString());
						}
						readmeEntry = archive.CreateEntry(relsFileName, CompressionLevel.Optimal);
						using( var relsFile = new StreamWriter(readmeEntry.Open()) )
						{
							relsFile.Write(excelTemplateRels, relTemplBuilder);
						}
						bwExportExcel.ReportProgress(100);
					}
				}
			}
			catch( Exception ex )
			{
				e.Result = ex;
			}
		}

		private void bwExportExcel_ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			progressBarControl.EditValue = e.ProgressPercentage;
		}
	}
}