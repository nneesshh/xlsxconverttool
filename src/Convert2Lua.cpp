#include "Convert2Lua.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#ifdef _WIN32
#include <io.h>
#else
#include <unistd.h>
#endif

#include <string>
#include <unordered_map>
#include <iostream>
#include <conio.h>

#include "libxl.h"

static char  stringSeparator = '\"';
static char *fieldSeparator = ",";
static char *lineSeparator = "\n";
static char *lineCompleteSeparator = ",";
static char *sentenceCompleteSeparator = "";

static std::unordered_map<std::string, bool> s_mapKeyExistence;

static bool check_key_exists(const std::string& sKey) {
	std::unordered_map<std::string, bool>::iterator iter = s_mapKeyExistence.find(sKey);
	if (iter == s_mapKeyExistence.end()) {
		s_mapKeyExistence[sKey] = true;
		return false;
	}
	return true;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
CConvert2Lua::CConvert2Lua(CXlsxConvertTool *app)
: _app(app) {

}

CConvert2Lua::~CConvert2Lua() {

}

int
CConvert2Lua::Convert(config_item_t *item) {
	char inputFileName[256];
	char outputFileName[256];
	TCHAR bookName[256] = { 0 };
	char sheetName[256] = { 0 };
	int sheetIdx;
	int titleRowId;
	int cellRow = 0, cellCol = 0;
	meta_table_t& meta_table = item->metaTable;

	libxl::Book *book;
	libxl::Sheet *sheet;
	bool bLoad;
	const char *errMsg;
	int sheetCount = 0, sheetType = 0;

	FILE*fp;
	//fprintf(fp, stderr, "DIR: %s\n\n", getcwd(NULL, 1024));

	sprintf(inputFileName, "%s", item->inputName);
	sprintf(outputFileName, "%s", item->outputName);
	sheetIdx = item->inputSheetIdx;
	titleRowId = item->inputTitleRowId;

	book = xlCreateXMLBook();
	if (!book) {
		fprintf(stderr, "libxl lib crashed!!!\n");
		return -1;
	}

#ifdef _UNICODE
	CXlsxConvertTool::C2w(inputFileName, bookName, 256);
	bLoad = book->load(bookName);
#else
	bLoad = book->load(inputFileName);
#endif
	errMsg = book->errorMessage();
	if (!bLoad || 0 != strcmp("ok", errMsg)) {
		fprintf(stderr, "file(%s) load failed, because %s!!!\n", inputFileName, errMsg);
		return -2;
	}

	// process sheet at index
	sheetCount = book->sheetCount();
	sheetType = book->sheetType(sheetIdx);
	if (sheetCount <= 0
		|| libxl::SHEETTYPE_SHEET != sheetType) {
		fprintf(stderr, "can't find sheet or sheet type error -- sheetIdx(%d), sheetCount(%d), sheetType(%d)!!!\n", sheetIdx, sheetCount, sheetType);
		return -3;
	}
	// open and parse the sheet
	sheet = book->getSheet(sheetIdx);

	// sheet name
#ifdef _UNICODE
		CXlsxConvertTool::W2c(sheet->name(), sheetName, 256);
#else
		sprintf(sheetName, sheet->name())£»
#endif

	// table name
	if (strlen(item->luaTableName) < 1) {
		sprintf(item->luaTableName, "%s", sheetName);
	}

	//
	fp = fopen(outputFileName, "ab+");
	if (fp == NULL) {
		fprintf(stderr, "create file failed : %s!!!\n", outputFileName);
		return -4;
	}

	//write UTF-8 BOM
	{
		//char bomheader[3] = { (char)0xEF, (char)0xBB, (char)0xBF }; // UTF - 8 BOM
		//fwrite(bomheader, sizeof(char), 3, fp);
	}

	int sheetLastRow = sheet->lastRow();
	int sheetFirstCol = meta_table.fromColumn;
	int sheetLastCol = meta_table.toColumn + 1;

	// find first all empty row as last row
	for (cellRow = 0; cellRow < sheetLastRow; ++cellRow) {
		bool bAllEmpty = true;
		// walk cells
		for (cellCol = sheetFirstCol; cellCol < sheetLastCol; ++cellCol) {
			libxl::CellType eCellType = sheet->cellType(cellRow, cellCol);
			if (libxl::CELLTYPE_BLANK != eCellType
				&& libxl::CELLTYPE_EMPTY != eCellType) {
				// check empty string
				if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
					const TCHAR *tstr = sheet->readStr(cellRow, cellCol);
					char buff[256] = { 0 };
					char *str = CXlsxConvertTool::W2c(tstr, buff, sizeof(buff));
#else
					const TCHAR *str = sheet->readStr(cellRow, cellCol);
#endif
					const TCHAR *sCell = sheet->readStr(cellRow, 0);
					int nLen = strlen(str);
					if (nLen > 0) {
						bAllEmpty = false;
						break;
					}
				}
				else {
					bAllEmpty = false;
					break;
				}
			}
		}

		if (bAllEmpty) {
			sheetLastRow = cellRow;
			break;
		}
	}

	// check data row exist or not
	if (sheetLastRow < item->inputTitleRowId + 1) {
		// table is empty
		// separate line
		fprintf(fp, "%s", lineSeparator);

		// start line -- table_name = {
		fprintf(fp, "--[[%s,%s]]\nlocal __tmp_tbl__ = {"
			, inputFileName
			, item->luaTableName);

		// close block
		fprintf(fp, "\n}%s;%s=%s or {};for k,v in pairs(__tmp_tbl__) do %s[k]=v end;%s"
			, sentenceCompleteSeparator
			, item->luaTableName
			, item->luaTableName
			, item->luaTableName
			, lineSeparator);

		//
		fclose(fp);
	}
	else if (sheetLastRow >= item->inputTitleRowId + 1) {
		//
		s_mapKeyExistence.clear();

		// at least exist one data row
		// process all rows of the sheet
		for (cellRow = 0; cellRow < sheetLastRow; ++cellRow) {
			const TCHAR *sCell = sheet->readStr(cellRow, 0);
			if (sCell && CXlsxConvertTool::StartsWith(sCell, (TCHAR *)"#")) {
				// ignore comment rows
				++meta_table.commentRowNum;
				continue;
			}

			// skip rows before title
			if (cellRow + 1 < titleRowId)
				continue;

			// walk cells
			for (cellCol = sheetFirstCol; cellCol < sheetLastCol; ++cellCol) {
				field_meta_t& meta = meta_table.arr[cellCol];

				// skip undefined field
				if (!meta.isField)
					continue;

				// collect title
				if (!meta_table.bTitleOk) {
					if (meta.isField) {
						libxl::CellType eCellType = sheet->cellType(cellRow, cellCol);
						if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
							const TCHAR *tstr = sheet->readStr(cellRow, cellCol);
							char buff[256] = { 0 };
							char *str = CXlsxConvertTool::W2c(tstr, buff, sizeof(buff));
#else
							const TCHAR *str = sheet->readStr(cellRow, cellCol);
#endif
							sprintf(meta.name, "%s", str);
						}
					}

					//
					if (meta.isToColumn) {
						cellCol = sheetLastCol;
						meta_table.bTitleOk = true;

						// separate line
						fprintf(fp, "%s", lineSeparator);

						// start line -- table_name = {
						fprintf(fp, "--[[%s,%s]]\nlocal __tmp_tbl__ = {"
							, inputFileName
							, item->luaTableName);
					}
				}
				else {
					// output cell
					OutputCell(fp, sheet, cellRow, cellCol, meta, inputFileName, sheetLastRow, sheetLastCol);
				}
			}
		}

		// close block
		fprintf(fp, "\n}%s;if not %s then %s =__tmp_tbl__;else for k,v in pairs(__tmp_tbl__) do %s[k]=v end;end;%s"
			, sentenceCompleteSeparator
			, item->luaTableName
			, item->luaTableName
			, item->luaTableName
			, lineSeparator);

		//
		fclose(fp);

		//
		book->release();
	}
	return 0;
}

// Output a string
void
CConvert2Lua::OutputCell(FILE *fp, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol) {
	libxl::Sheet *sheet_ = static_cast<libxl::Sheet *>(sheet);

	// cell type
	libxl::CellType eCellType = sheet_->cellType(row, col);

	// process none visible and empty
	if (libxl::CELLTYPE_BLANK == eCellType
		|| libxl::CELLTYPE_EMPTY == eCellType
		|| libxl::CELLTYPE_ERROR == eCellType
		|| libxl::SHEETSTATE_VISIBLE != sheet_->hidden()) {
		// output empty string
		OutputString(fp, "", meta, inputFileName);

		//
		if (meta.isToColumn && row + 1 < lastRow)
			fprintf(fp, "%s", lineCompleteSeparator);
	}
	else if (libxl::CELLTYPE_NUMBER == eCellType) {
		double number = sheet_->readNum(row, col);
		OutputNumber(fp, number, meta, inputFileName);

		//
		if (meta.isToColumn && row + 1 < lastRow)
			fprintf(fp, "%s", lineCompleteSeparator);
	}
	else if (libxl::CELLTYPE_STRING == eCellType) {
#ifdef _UNICODE
		const TCHAR *tstr = sheet_->readStr(row, col);
		char buff[1024 * 16] = { 0 };
		char *str = CXlsxConvertTool::W2c(tstr, buff, sizeof(buff));
#else
		const TCHAR *str = sheet_->readStr(cellRow, cellCol);
#endif
		OutputString(fp, str, meta, inputFileName);

		//
		if (meta.isToColumn && row + 1 < lastRow)
			fprintf(fp, "%s", lineCompleteSeparator);
	}
	else if (libxl::CELLTYPE_BOOLEAN == eCellType) {
		bool b = sheet_->readBool(row, col);
		//OutputString(fp, b ? "true" : "false", meta, inputFileName);
	}
	else {
		//OutputString(fp, "", meta, inputFileName);
	}
}

// Output a string
void
CConvert2Lua::OutputString(FILE *fp, const char *str, field_meta_t& meta, const char *inputFileName) {
	if (meta.isFromColumn) {
		// check key exists
		bool bExist = check_key_exists(str);
		if (bExist) {
			fprintf(stderr, "[Convert2Lua] WARNING: key(\"%s\")val(\"%s\") is duplicated -- file(%s)!!!\n",
				meta.name, str, inputFileName);
		}
		fprintf(fp, "%s\t%s = {%s=%c%s%c", lineSeparator, str, meta.name, stringSeparator, str, stringSeparator);
	}
	else {
		// skip empty string
//		if (strlen(str) > 0) {
			fprintf(fp, "%s", fieldSeparator);
			fprintf(fp, "%s=%c%s%c", meta.name, stringSeparator, str, stringSeparator);
//		}
	}

	if (meta.isToColumn) {
		fprintf(fp, "}");
	}
}

// Output a number
void
CConvert2Lua::OutputNumber(FILE *fp, const double number, field_meta_t& meta, const char *inputFileName) {
	if (meta.isFromColumn) {
		// check key exists
		std::string sStr = std::to_string(number);
		bool bExist = check_key_exists(sStr);
		if (bExist) {
			fprintf(stderr, "[Convert2Lua] WARNING: key(\"%s\")val(\"%.15g\") is duplicated -- file(%s)!!!\n",
				meta.name, number, inputFileName);
		}
		fprintf(fp, "%s\t[%.15g] = {%s=%.15g", lineSeparator, number, meta.name, number);
	}
	else {
		fprintf(fp, "%s", fieldSeparator);
		fprintf(fp, "%s=%.15g", meta.name, number);
	}

	if (meta.isToColumn) {
		fprintf(fp, "}");
	}
}