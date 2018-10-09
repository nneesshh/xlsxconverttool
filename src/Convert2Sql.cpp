#include "Convert2Sql.h"

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#ifdef _WIN32
#include <io.h>
#else
#include <unistd.h>
#endif

#include <iostream>
#include <conio.h>

#include "libxl.h"

static char  stringSeparator = '\'';
static char *fieldSeparator = ",";
static char *lineSeparator = "\n";
static char *lineCompleteSeparator = ",";
static char *sentenceCompleteSeparator = ";";

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
CConvert2Sql::CConvert2Sql(CXlsxConvertTool *app)
: _app(app) {

}

CConvert2Sql::~CConvert2Sql() {

}

int
CConvert2Sql::Convert(config_item_t *item) {
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
	if (strlen(item->sqlTableName) < 1) {
		sprintf(item->sqlTableName, "%s", sheetName);
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
		if (strlen(item->deleteClause) >= 3) {
			// separate line
			fprintf(fp, "%s", lineSeparator);

			// start line -- DELETE FROM xxxx
			fprintf(fp, "DELETE FROM `%s` WHERE %s;"
				, item->sqlTableName
				, item->deleteClause);
		}

		// close block
		fprintf(fp, "%s", lineSeparator);

		//
		fclose(fp);
	}
	else {
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

						if (strlen(item->deleteClause) >= 3) {
							// separate line
							fprintf(fp, "%s", lineSeparator);

							// start line -- DELETE FROM xxxx
							fprintf(fp, "DELETE FROM `%s` WHERE %s;"
								, item->sqlTableName
								, item->deleteClause);
						}

						// separate line
						fprintf(fp, "%s", lineSeparator);

						// start line -- INSERT INTO xxx
						fprintf(fp, "INSERT INTO `%s` (", item->sqlTableName);
						{
							bool bFirst = true;
							int i;
							for (i = sheetFirstCol; i < sheetLastCol; ++i) {
								field_meta_t& tmp = meta_table.arr[i];
								if (tmp.isField) {
									if (bFirst) {
										fprintf(fp, "`%s`", tmp.name);
										bFirst = false;
									}
									else {
										// separate line
										fprintf(fp, "%s ", fieldSeparator);

										// field name
										fprintf(fp, "`%s`", tmp.name);
									}
								}

								if (tmp.isToColumn)
									break;
							}
						}
						fprintf(fp, ") %sVALUES ", lineSeparator);
					}
				}
				else {
					// output cell
					OutputCell(fp, sheet, cellRow, cellCol, meta, inputFileName, sheetLastRow, sheetLastCol);
				}
			}
		}

		// tail
		fprintf(fp, "%sON DUPLICATE KEY UPDATE ", lineSeparator);

		bool bFirst = true;
		int i;
		for (i = sheetFirstCol; i < sheetLastCol; ++i) {
			field_meta_t& tmp = meta_table.arr[i];

			if (tmp.isField && !tmp.isKey) {
				if (bFirst) {
					fprintf(fp, "`%s` = VALUES(`%s`)", tmp.name, tmp.name);
					bFirst = false;
				}
				else {
					fprintf(fp, ", `%s` = VALUES(`%s`)", tmp.name, tmp.name);
				}
			}
		}

		// close block
		fprintf(fp, "%s%s", sentenceCompleteSeparator, lineSeparator);

		//
		fclose(fp);

		//
		book->release();
	}
	return 0;
}

// Output a string
void
CConvert2Sql::OutputCell(FILE *fp, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol) {
	libxl::Sheet *sheet_ = static_cast<libxl::Sheet *>(sheet);

	// cell type
	libxl::CellType eCellType = sheet_->cellType(row, col);

	// process none visible and empty
	if (libxl::CELLTYPE_BLANK == eCellType
		|| libxl::CELLTYPE_EMPTY == eCellType
		|| libxl::CELLTYPE_ERROR == eCellType
		|| libxl::SHEETSTATE_VISIBLE != sheet_->hidden()) {
		// output empty string
		OutputString(fp, "", meta);

		//
		if (meta.isToColumn && row + 1 < lastRow)
			fprintf(fp, "%s", lineCompleteSeparator);
	}
	else if (libxl::CELLTYPE_ERROR == eCellType) {
		FILE *ferr = fopen("xlsxconverttool.error", "at+");
		fprintf(ferr, "<%s> has error cell -- row(%d)col(%d)\n", inputFileName, row, col);
		fclose(ferr);
	}
	else if (libxl::CELLTYPE_NUMBER == eCellType) {
		double number = sheet_->readNum(row, col);
		OutputNumber(fp, number, meta);

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
		OutputString(fp, str, meta);

		//
		if (meta.isToColumn && row + 1 < lastRow)
			fprintf(fp, "%s", lineCompleteSeparator);
	}
	else if (libxl::CELLTYPE_BOOLEAN == eCellType) {
		bool b = sheet_->readBool(row, col);
		//OutputString(fp, b ? "true" : "false", meta);
	}
	else {
		//OutputString(fp, "", meta);
	}
}

// Output a string
void
CConvert2Sql::OutputString(FILE *fp, const char *str, field_meta_t& meta) {
	if (meta.isFromColumn) {
		fprintf(fp, "(%c%s%c", stringSeparator, str, stringSeparator);
	}
	else {
		fprintf(fp, "%s", fieldSeparator);
		fprintf(fp, "%c%s%c", stringSeparator, str, stringSeparator);
	}

	if (meta.isToColumn) {
		fprintf(fp, ")");
	}
}

// Output a number
void
CConvert2Sql::OutputNumber(FILE *fp, const double number, field_meta_t& meta) {
	if (meta.isFromColumn) {
		fprintf(fp, "%s(%.15g", lineSeparator, number);
	}
	else {
		fprintf(fp, "%s", fieldSeparator);
		fprintf(fp, "%.15g", number);
	}

	if (meta.isToColumn) {
		fprintf(fp, ")");
	}
}