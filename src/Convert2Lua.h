#ifndef CONVERT2LUA_H
#define CONVERT2LUA_H

#include "Config.h"

#include "XlsxConvertTool.h"

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
class CConvert2Lua {
public:
	CConvert2Lua(CXlsxConvertTool *app);
	virtual ~CConvert2Lua();

	int Convert(config_item_t *item);

public:
	static void OutputCell(FILE *fp, void *sheet, int row, int col, field_meta_t& meta, const char *inputFileName, int lastRow, int lastCol);
	static void OutputString(FILE *fp, const char *str, field_meta_t& meta, const char *inputFileName);
	static void OutputNumber(FILE *fp, const double number, field_meta_t& meta, const char *inputFileName);

public:
	CXlsxConvertTool *_app;
};

#endif
