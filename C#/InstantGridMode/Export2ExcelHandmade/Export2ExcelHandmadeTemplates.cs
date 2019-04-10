namespace InstanGridMode
{
	public partial class Export2ExcelHandmade : DevExpress.XtraEditors.XtraForm
	{
		/// <summary>
		/// Rels file
		/// </summary>
		const string excelTemplateRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id =\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>{0}</Relationships>";
		const string excelTemplateRelsForWorksheet = "<Relationship Id = \"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{1}.xml\"/>";
		const string excelTemplateGeneralRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
		/// <summary>
		/// Styles file
		/// </summary>
		const string excelTemplateStyles = "<?xml version=\"1.0\" encoding=\"utf-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"{0}\" /></numFmts><fonts count=\"2\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font><font><b /><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills><borders count=\"2\"><border><left /><right /><top /><bottom /><diagonal /></border><border><left style=\"thin\" /><right style=\"thin\" /><top style=\"thin\" /><bottom style=\"thin\" /><diagonal /></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs><cellXfs count=\"5\"><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"1\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"164\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"1\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles><dxfs count=\"0\" /></styleSheet>";
		const string excelTemplateStylesDateFormat = "DD.MM.YYYY hh:mm:ss";
		/// <summary>
		/// Shared string file
		/// </summary>
		const string excelTemplateSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"88\" uniqueCount=\"88\"></sst>";
		/// <summary>
		/// Workbook file
		/// </summary>
		const string excelTemplateWorkbook = "<?xml version=\"1.0\" encoding=\"utf-8\"?><workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><bookViews><workbookView /></bookViews><sheets>{0}</sheets><definedNames>{1}</definedNames><calcPr fullCalcOnLoad=\"1\" /></workbook>";
		const string excelTemplateWorkbookSheet = "<sheet name=\"{0}\" sheetId=\"{1}\" r:id=\"rId{2}\" />";
		const string excelTemplateWorkbookSheetDefName = "<definedName name =\"_xlnm._FilterDatabase\" localSheetId=\"{0}\" hidden=\"1\">'{1}'!$A$1:${2}${3}</definedName>";
		/// <summary>
		/// Worksheet file
		/// </summary>
		const string excelTemplateWorksheetHeader = "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:{0}{1}\" /><sheetViews><sheetView workbookViewId=\"{2}\"><pane ySplit=\"1\" topLeftCell=\"A2\" state=\"frozen\" activePane=\"bottomLeft\" /><selection pane=\"bottomLeft\" activeCell=\"A1\" sqref=\"A1\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\" /><cols>";
		const string excelTemplateWorksheetColumn = "<col min=\"1\" max=\"{0}\" width=\"20\" customWidth=\"1\"/>";
		const string excelTemplateWorksheetBeginBody = "</cols><sheetData>";
		const string excelTemplateWorksheetRow = "<row r=\"{0}\">{1}</row>";
		const string excelTemplateWorksheetCol = "<c r =\"{0}{1}\" s=\"{2}\" {3}>{4}</c>"; // 3 -> t=\"inlineStr\"
		const string excelTemplateWorksheetEndBody = "</sheetData><autoFilter ref=\"A1:{0}{1}\" /><headerFooter /></worksheet>";
		/// <summary>
		/// Content types file
		/// </summary>
		const string excelTemplateContentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default ContentType=\"application/xml\" Extension=\"xml\"/><Default ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" Extension=\"rels\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" PartName=\"/xl/workbook.xml\" />{0}<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" PartName=\"/xl/styles.xml\" /><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" PartName=\"/xl/sharedStrings.xml\" /></Types>";
		const string excelTemplateContentTypesWorkSheetOverride = "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet{0}.xml\" />";
    }
}