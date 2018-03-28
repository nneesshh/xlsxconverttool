#ifndef CONVERT2SQL_H
#define CONVERT2SQL_H

#include "Config.h"

#include "XlsxConvertTool.h"

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
class CConvert2Sql {
public:
	CConvert2Sql(CXlsxConvertTool *app);
	virtual ~CConvert2Sql();

	int Convert(config_item_t *item);

public:
	static void OutputCell(FILE *fp, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol);
	static void OutputString(FILE *fp, const char *str, field_meta_t& meta);
	static void OutputNumber(FILE *fp, const double number, field_meta_t& meta);

public:
	CXlsxConvertTool *_app;
};

#endif
