# -*- coding: UTF-8 -*-
import argparse
import os
from datetime import datetime
import traceback
import time
import re
import json
import sqlite3
import zipfile
from html import escape

excelTemplateRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/><Relationship Id =\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>{0}</Relationships>"
excelTemplateRelsForWorksheet = "<Relationship Id = \"rId{0}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{1}.xml\"/>"
excelTemplateGeneralRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
excelTemplateStyles = "<?xml version=\"1.0\" encoding=\"utf-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"{0}\" /></numFmts><fonts count=\"2\"><font><sz val=\"11\" /><name val=\"Calibri\" /></font><font><b /><sz val=\"11\" /><name val=\"Calibri\" /></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\" /></fill><fill><patternFill patternType=\"gray125\" /></fill></fills><borders count=\"2\"><border><left /><right /><top /><bottom /><diagonal /></border><border><left style=\"thin\" /><right style=\"thin\" /><top style=\"thin\" /><bottom style=\"thin\" /><diagonal /></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" /></cellStyleXfs><cellXfs count=\"5\"><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"0\" applyNumberFormat=\"1\" fontId=\"1\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"164\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /><xf numFmtId=\"1\" applyNumberFormat=\"1\" fontId=\"0\" applyFont=\"1\" borderId=\"1\" applyBorder=\"1\" xfId=\"0\" applyProtection=\"1\" /></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\" /></cellStyles><dxfs count=\"0\" /></styleSheet>"
excelTemplateStylesDateFormat = "DD.MM.YYYY hh:mm:ss"
excelTemplateSharedString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"88\" uniqueCount=\"88\"></sst>"
excelTemplateWorkbook = "<?xml version=\"1.0\" encoding=\"utf-8\"?><workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><bookViews><workbookView /></bookViews><sheets>{0}</sheets><definedNames>{1}</definedNames><calcPr fullCalcOnLoad=\"1\" /></workbook>"
excelTemplateWorkbookSheet = "<sheet name=\"{0}\" sheetId=\"{1}\" r:id=\"rId{2}\" />"
excelTemplateWorkbookSheetDefName = "<definedName name =\"_xlnm._FilterDatabase\" localSheetId=\"{0}\" hidden=\"1\">'{1}'!$A$1:${2}${3}</definedName>"
excelTemplateWorksheetHeader = "<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:{0}{1}\" /><sheetViews><sheetView workbookViewId=\"{2}\"><pane ySplit=\"1\" topLeftCell=\"A2\" state=\"frozen\" activePane=\"bottomLeft\" /><selection pane=\"bottomLeft\" activeCell=\"A1\" sqref=\"A1\" /></sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\" /><cols>"
excelTemplateWorksheetColumn = "<col min=\"1\" max=\"{0}\" width=\"20\" customWidth=\"1\"/>"
excelTemplateWorksheetBeginBody = "</cols><sheetData>"
excelTemplateWorksheetRow = "<row r=\"{0}\">{1}</row>"
excelTemplateWorksheetCol = "<c r =\"{0}{1}\" s=\"{2}\" {3}>{4}</c>" #// 3 -> t=\"inlineStr\"
excelTemplateWorksheetEndBody = "</sheetData><autoFilter ref=\"A1:{0}{1}\" /><headerFooter /></worksheet>"
excelTemplateContentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default ContentType=\"application/xml\" Extension=\"xml\"/><Default ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" Extension=\"rels\"/><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" PartName=\"/xl/workbook.xml\" />{0}<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" PartName=\"/xl/styles.xml\" /><Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" PartName=\"/xl/sharedStrings.xml\" /></Types>"
excelTemplateContentTypesWorkSheetOverride = "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet{0}.xml\" />"

# CONSOLE LOG
cfromat = "[{0}] {1}{2}"

def print_message(message, level=0, end="\n"):
    level_indent = " " * level
    try:
        print(cfromat.format(datetime.now(), level_indent, message), end=end)
    except KeyboardInterrupt:
        print('Pressed CTRL+C')
        os._exit(-1)
    except:
        print('programmers did not fix encoding))')

        
class SQLiteDataProvider:
    def __init__(self, db_path, params, header_fname, print_log=False):
        """ Initialize provider obj """
        self._init = False
        self._header_file_delimiter = ">"
        
        self.db_path = db_path
        self.header_fname = header_fname
        self.params = json.loads(params)
        self.TABLENAME = "calls_v"
        self.print_log = print_log
        self.header_dict = None
        self.rows_count = None

    def __enter__(self):
        """ Enter in 'with' """
        self.init_provider()
        return self

    def init_provider(self):
        """ Init SQLite data provider """
        if self.print_log: print_message("Initialize SQLite data provider")
        if self.print_log: print_message("Connecting to database '{}'".format(self.db_path), 2)
        self._dbconn = sqlite3.connect(self.db_path)
        self._cur = self._dbconn.cursor()
        if self.print_log: print_message("Check columns...", 2)
        if not self.params.get("column_list"):
            raise Exception("Empty list with columns!")
        if self.print_log: print_message("Prcoess headers of columns...", 2)
        self.header_dict = dict()
        for col_descr_row in open(self.header_fname, "r", encoding="utf-8"):
            tmp_descr = list(map(lambda s: s.strip(), 
                                 col_descr_row.split(self._header_file_delimiter)))
            if len(tmp_descr) < 2:
                raise Exception("Error in file with columns desrciptions: columns in description less then 2!")
            elif len(tmp_descr) == 2:
                self.header_dict[tmp_descr[0]] = [tmp_descr[1], "str"]
            else:
                self.header_dict[tmp_descr[0]] = tmp_descr[1:3]
        sql = """SELECT count(*) FROM {} {}""".format(
            self.TABLENAME,
            "WHERE {}".format(self.params.get("where")) \
                if self.params.get("where") else "")
        if self.print_log: print_message("Get rows count from result... ", 2, end="")
        self._cur.execute(sql)
        self.rows_count = self._cur.fetchone()[0]
        if self.print_log: print(self.rows_count)
        sql = sql.replace("count(*)", ",".join(self.params.get("column_list")))
        if self.print_log: print_message("Execute SQL: '{}'".format(sql.strip()), 2)
        self._cur.execute(sql)
        self._init = True
        if self.print_log: print_message("Initialize complete!", 2)

    def get_columns(self):
        """ Return list of pairs (column name, column type) """
        if not self._init:
            raise Exception("Provider is not initialized!")
        return [self.header_dict.get(col) \
                for col in self.params.get("column_list")]

    def get_next_row(self):
        """ Return next row from SQL request """
        if not self._init:
            raise Exception("Provider is not initialized!")
        if self._cur:
            return self._cur.fetchone()

    def __exit__(self, exc_type, exc_value, traceback):
        """ Dispose object """
        if self._dbconn:
            self._dbconn.close()
            
    def __str__(self):
        """ String represetnation """
        return "SQLiteDataProvider(init={}, db={})".format(self._init, self.db_path)
       
       
def number_to_letters(number):
    result = ""
    if number > 0:
        alphabets = (number - 1) // 26
        remainder = (number - 1) % 26
        result = (chr(ord('A') + remainder))
        if alphabets > 0:
            result = number_to_letters(alphabets) + result
    else:
        result = ""
    return result

    
def datetime2ole(date):
    date = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
    OLE_TIME_ZERO = datetime(1899, 12, 30)
    delta = date - OLE_TIME_ZERO
    return float(delta.days) + (float(delta.seconds) / 86400)

    
def export_to_excel(data_provider, output_data, max_rows_in_file=500000, verbose_log=True):
    with data_provider as provider:
        print_message("Export data from data provider {} to excel file '{}'".format(
            data_provider, output_data))
        if os.path.isfile(output_data):
            os.remove(output_data)
        with zipfile.ZipFile(output_data, "a", zipfile.ZIP_DEFLATED) as excel_file:
            if verbose_log: print_message("Create xlsx structure", 2)
            if verbose_log: print_message("Create styles file", 4)
            with excel_file.open(r"xl\styles.xml", "w") as elem_file:
                elem_file.write(bytes(excelTemplateStyles.format(
                    excelTemplateStylesDateFormat), 'utf-8'))
            if verbose_log: print_message("Create general rels file", 4)
            with excel_file.open(r"_rels\.rels", "w") as elem_file:
                elem_file.write(bytes(excelTemplateGeneralRels, 'utf-8'))
            if verbose_log: print_message("Create shared strings file", 4)
            with excel_file.open(r"xl\sharedStrings.xml", "w") as elem_file:
                elem_file.write(bytes(excelTemplateSharedString, 'utf-8'))
            print_message("Export data...", 2)
            wb_sheets_templ_builder = ""
            wb_sheets_templ_sheet_def_name_builder = ""
            ct_work_sheet_override_builder = ""
            rel_templ_builder = ""
            data_columns = provider.get_columns()
            col_count = len(data_columns)
            worksheets_count = provider.rows_count // max_rows_in_file + 1
            for worksheet_index in range(worksheets_count):
                if verbose_log: print_message("Create worksheet #{} (total{})".format(
                    worksheet_index + 1,
                    worksheets_count
                ), 4)
                ws_fname = r"xl\worksheets\sheet{}.xml".format(worksheet_index + 1)
                ws_name ="Billing #{}".format(worksheet_index + 1) if worksheets_count > 1 else "Billing"
                with excel_file.open(ws_fname, "w") as ws_file:
                    start_row_in_grid = worksheet_index * max_rows_in_file
                    rows_in_current_file = min(provider.rows_count - start_row_in_grid, max_rows_in_file)
                    end_columns_range = number_to_letters(col_count)
                    # Add in builder cur sheet
                    wb_sheets_templ_builder += excelTemplateWorkbookSheet.format(ws_name, worksheet_index + 1, worksheet_index + 3)
                    wb_sheets_templ_sheet_def_name_builder += excelTemplateWorkbookSheetDefName.format(worksheet_index, ws_name, 
                        end_columns_range, rows_in_current_file + 1)
                    ct_work_sheet_override_builder += excelTemplateContentTypesWorkSheetOverride.format(worksheet_index + 1)
                    rel_templ_builder += excelTemplateRelsForWorksheet.format(worksheet_index + 3, worksheet_index + 1)
                    # Create header, coldata and open body in worksheet
                    ws_file.write(bytes(excelTemplateWorksheetHeader.format(end_columns_range, rows_in_current_file + 1, 0), 'utf-8')) 
                    ws_file.write(bytes(excelTemplateWorksheetColumn.format(col_count), 'utf-8')) 
                    ws_file.write(bytes(excelTemplateWorksheetBeginBody, 'utf-8'))
                    # Write header
                    row_builder = ""
                    for num, col_data in enumerate(data_columns):
                        column_name, column_type = col_data
                        row_builder += excelTemplateWorksheetCol.format(number_to_letters(num + 1), 1, 2,
                            "t =\"inlineStr\"", "<is><t>{}</t></is>".format(escape(column_name)))
                    ws_file.write(bytes(excelTemplateWorksheetRow.format(1, row_builder), 'utf-8'))
                    #  Export data rows
                    cur_progress = 0;
                    if verbose_log: print_message("Export rows #{}-#{} (total {}) from grid to file {}".format(
                        start_row_in_grid + 1,
                        start_row_in_grid + rows_in_current_file,
                        rows_in_current_file,
                        ws_fname
                    ), 5)
                    for i in range(rows_in_current_file):
                        res = provider.get_next_row()
                        if not res:
                            break
                        row_builder = ""
                        for j, value in enumerate(res):
                            value_type = data_columns[j][1]
                            is_date = "date" == value_type
                            is_int = "int" == value_type
                            use_inline_str = not is_date and not is_int
                            row_builder += excelTemplateWorksheetCol.format(
                                number_to_letters(j + 1), i + 2,
                                3 if is_date else
                                    4 if is_int else 1,
                                "t =\"inlineStr\"" if use_inline_str else "",
                                "<is><t>{}</t></is>".format("" if value == None else escape(value)) if use_inline_str 
                                else "<v>{}</v>".format("" if value == None else datetime2ole(value) if is_date else value)
                            )
                            ws_file.write(bytes(excelTemplateWorksheetRow.format(i + 2, row_builder), 'utf-8'))
                    ws_file.write(bytes(excelTemplateWorksheetEndBody.format(end_columns_range, rows_in_current_file + 1), 'utf-8'))
            if verbose_log: print_message("Create content file", 4)
            with excel_file.open(r"[Content_Types].xml", "w") as elem_file:
                elem_file.write(bytes(excelTemplateContentTypes.format(ct_work_sheet_override_builder), 'utf-8'))
            if verbose_log: print_message("Create workbook file", 4)
            with excel_file.open(r"xl\workbook.xml", "w") as elem_file:
                elem_file.write(bytes(excelTemplateWorkbook.format(wb_sheets_templ_builder,
                    wb_sheets_templ_sheet_def_name_builder), 'utf-8'))
            if verbose_log: print_message("Create rels file", 4)
            with excel_file.open(r"xl\_rels\workbook.xml.rels", "w") as elem_file:
                elem_file.write(bytes(excelTemplateRels.format(rel_templ_builder), 'utf-8'))
            print_message("Complete export to excel!", 2)


try:
    _parser = argparse.ArgumentParser()
    requiredNamed = _parser.add_argument_group('Required arguments')
    requiredNamed.add_argument("-d", action="store", dest="DATABASE_PATH", help="Database path", type=str, required=True)
    requiredNamed.add_argument("-o", action="store", dest="OUTPUT_NAME", help="Output excel file path", type=str, required=True)
    requiredNamed.add_argument("-p", action="store", dest="PARAMS", help="Parameters", type=str, required=True)
    requiredNamed.add_argument("-c", action="store", dest="HEADER_FILE", help="Header file path", type=str, required=True)
    requiredNamed.add_argument("-v", action="store_true", dest="VERBOSE", help="Verbose")

    _command_args = _parser.parse_args()
    _db_path = _command_args.DATABASE_PATH
    _output_excel_fname = _command_args.OUTPUT_NAME
    _params = _command_args.PARAMS
    _header_fname = _command_args.HEADER_FILE
    _verbose_mode = _command_args.VERBOSE

    start_time = datetime.now()
    export_to_excel(SQLiteDataProvider(_db_path,  _params, _header_fname, _verbose_mode), _output_excel_fname, verbose_log=_verbose_mode)
    os._exit(0)
except KeyboardInterrupt:
    print('Pressed CTRL+C')
    os._exit(-1)
except:
    print(traceback.format_exc())
    os._exit(-1)
