#include "XlsxConvertTool.h"

#include <stdio.h>
#include <stdlib.h>
#include <algorithm>


#include "pugixml/pugiconfig.hpp"
#include "pugixml/pugixml.hpp"

#include "Convert2Sql.h"
#include "Convert2Lua.h"

static int 
column_to_num(const char *sColumn) {
	if (NULL == sColumn) {
		return -1;
	}
	if (0 == strlen(sColumn))
		return -2;

	const char *ptr;
	char ch;
	unsigned int i;

	ptr = sColumn;
	for (i = 0; i < strlen(sColumn); ++i) {
		ch = *ptr;
		if (ch < 'A' && ch > 'Z') {
			return -3;
		}
		++ptr;
	}

	//
	int result = -1;
	ptr = sColumn;
	for (i = 0; i < strlen(sColumn); ++i) {
		char ch = *ptr;
		if (ch >= 'A' && ch <= 'Z') {
			result = (result + 1) * 26 + (ch - 'A');
		}
		++ptr;
	}
	return result;
}

int
split(const char *str, int str_len, char **av, int av_max, char c) {
	int i, j;
	char *ptr = (char*)str;
	int count = 0;

	if (!str_len) str_len = strlen(ptr);

	for (i = 0, j = 0; i < str_len&& count < av_max; ++i) {
		if (ptr[i] != c)
			continue;
		else
			ptr[i] = 0x0;

		av[count++] = &(ptr[j]);
		j = i + 1;
		continue;
	}

	if (j < i) av[count++] = &(ptr[j]);

	return count;
}

int
split2d(const char *str, int str_len, char *(*av)[32], int av_max, char c1, char c2, int *out_arr_n2) {
	char* arr1[256];
	int i, n1, n2;

	int av1_max = std::min<int>(av_max, 256);
	int av2_max = 32;

	n1 = split(str, str_len, arr1, av1_max, c1);
	for (i = 0;i < n1; ++i) {
		n2 = split(arr1[i], strlen(arr1[i]), av[i], av2_max, c2);
		out_arr_n2[i] = n2;
	}
	return n1;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
CXlsxConvertTool::CXlsxConvertTool() {

}

CXlsxConvertTool::~CXlsxConvertTool() {

}

void
CXlsxConvertTool::SetupMetaTable() {
	// truncate error file
	{
		FILE *ferr = fopen("xlsxconverttool.error", "wt+");
		fclose(ferr);
	}

	// parse
#define FILE_BUF_SIZE 1024 * 512
	char *filebuf = new char[FILE_BUF_SIZE];
	size_t file_size = 0;
	char path[256] = { 0 };
	FILE *fp;
	
	sprintf(path, "%s", "xlsxconverttool.xml");

	fp = fopen(path, "r");
	if (nullptr == fp) {
		fprintf(stderr, "file <%s> open failed!!!\n", path);

		FILE *ferr = fopen("xlsxconverttool.error", "at+");
		fprintf(ferr, "file <%s> open failed!!!\n", path);
		fclose(ferr);
		exit(-1);
	}
	else {
		file_size = fread(filebuf, 1, FILE_BUF_SIZE, fp);
		if (0 == file_size) {
			fprintf(stderr, "file <%s> is empty!!!\n", path);

			FILE *ferr = fopen("xlsxconverttool.error", "at+");
			fprintf(ferr, "file <%s> is empty!!!\n", path);
			fclose(ferr);
			exit(-1);
		}
		else if (file_size == FILE_BUF_SIZE) {
			fprintf(stderr, "file <%s> is empty!!!\n", path);

			FILE *ferr = fopen("xlsxconverttool.error", "at+");
			fprintf(ferr, "file <%s> is too big -- (%d/%d)bytes!!!\n", path, file_size, FILE_BUF_SIZE - 1);
			fclose(ferr);
			exit(-1);
		}
		if (file_size > 0) {
			pugi::xml_document doc;
			pugi::xml_parse_result result = doc.load_buffer(filebuf, file_size);
			pugi::xml_node root_children = doc.child("Root");
			for (pugi::xml_node_iterator it = root_children.begin(); it != root_children.end(); ++it) {
				std::string name = it->name();

				//
				if (name == "Sql") {
					for (pugi::xml_node_iterator it2 = it->begin(); it2 != it->end(); ++it2) {
						std::string strInput = it2->attribute("InputName").value();
						if (strInput.length() <= 0)
							continue;

						int nSheetIdx = atoi(it2->attribute("SheetIdx").value());
						int nTitleRowId = atoi(it2->attribute("TitleRowId").value());
						std::string strFields = it2->attribute("Fields").value();
						std::string strKeys = it2->attribute("Keys").value();
						std::string strOutput = it2->attribute("OutputName").value();
						std::string strDeleteClause = it2->attribute("DeleteClause").value();

						config_item_t item;
						meta_table_t& metaTable = item.metaTable;
						memset(&metaTable, 0, sizeof(metaTable));

						metaTable.bTitleOk = false;
						metaTable.commentRowNum = 0;

						// init
						{
							int i;
							for (i = 0; i < sizeof(metaTable.arr) / sizeof(metaTable.arr[0]); ++i) {
								field_meta_t& meta = metaTable.arr[i];
								meta.id = i + 1;
							}
						}

						item.c = CONVERT_2_SQL;
						sprintf(item.inputName, "%s", strInput.c_str());
						item.inputSheetIdx = nSheetIdx;
						item.inputTitleRowId = nTitleRowId;
						sprintf(item.inputFields, "%s", strFields.c_str());
						sprintf(item.inputKeys, "%s", strKeys.c_str());
						sprintf(item.outputName, "%s", strOutput.c_str());
						sprintf(item.deleteClause, "%s", strDeleteClause.c_str());

						// fields format: "A, C-Z,..."
						{
							char * chArr[8][32] = { 0 };
							int arrN2[8] = { 0 };
							int n = split2d(strFields.c_str(), strFields.length(), chArr, 8, ',', '-', arrN2);
							int i, j;
							int nFromColumn = -1;
							int nToColumn = -1;
							for (i = 0; i < n; ++i) {
								int nField1 = column_to_num(chArr[i][0]);
								int nField2 = column_to_num(chArr[i][1]);
								nField2 = nField2 > 0 ? nField2 : nField1;

								for (j = nField1; j <= nField2; ++j) {
									metaTable.arr[j].isField = true;

									if (-1 == nFromColumn) {
										nFromColumn = j;
									}

									//
									nToColumn = j;
								}
							}

							// fields : from - to
							metaTable.fromColumn = nFromColumn;
							metaTable.toColumn = nToColumn;

							if (nFromColumn >= 0) {
								metaTable.fromColumn = nFromColumn;
								metaTable.arr[nFromColumn].isFromColumn = true;

							}

							if (nToColumn >= 0) {
								metaTable.toColumn = nToColumn;
								metaTable.arr[nToColumn].isToColumn = true;
							}
						}

						// fields format: "A, C-Z,..."
						{
							char * chArr[8][32] = { 0 };
							int arrN2[8] = { 0 };
							int n = split2d(strKeys.c_str(), strKeys.length(), chArr, 8, ',', '-', arrN2);
							int i, j;
							for (i = 0; i < n; ++i) {
								int nKey1 = column_to_num(chArr[i][0]);
								int nKey2 = column_to_num(chArr[i][1]);
								nKey2 = nKey2 > 0 ? nKey2 : nKey1;

								for (j = nKey1; j <= nKey2; ++j) {
									if (metaTable.arr[j].isField)
										metaTable.arr[j].isKey = true;
								}
							}
						}

						//
						_config.push_back(item);

						// truncate output file 
						FILE *fp_to_truncate = nullptr;
						if (strlen(item.outputName) > 1) {
							fp_to_truncate = fopen(item.outputName, "wb+");
						}

						if (fp_to_truncate) {
							fclose(fp_to_truncate);
						}
						else {
							fprintf(stderr, "output file <%s> truncate failed!!!\n", item.outputName);

							FILE *ferr = fopen("xlsxconverttool.error", "at+");
							fprintf(ferr, "output file <%s> truncate failed!!!\n", item.outputName);
							fclose(ferr);
							exit(-1);
						}
					}
				}
				else if (name == "Lua") {
					for (pugi::xml_node_iterator it2 = it->begin(); it2 != it->end(); ++it2) {
						std::string strInput = it2->attribute("InputName").value();
						if (strInput.length() <= 0)
							continue;

						int nSheetIdx = atoi(it2->attribute("SheetIdx").value());
						int nTitleRowId = atoi(it2->attribute("TitleRowId").value());
						std::string strFields = it2->attribute("Fields").value();
						std::string strKeys = "";
						std::string strOutput = it2->attribute("OutputName").value();
						std::string strLuaTable = it2->attribute("LuaTableName").value();

						config_item_t item;
						meta_table_t& metaTable = item.metaTable;
						memset(&metaTable, 0, sizeof(metaTable));

						metaTable.bTitleOk = false;
						metaTable.commentRowNum = 0;

						// init
						{
							int i;
							for (i = 0; i < sizeof(metaTable.arr) / sizeof(metaTable.arr[0]); ++i) {
								field_meta_t& meta = metaTable.arr[i];
								meta.id = i + 1;
							}
						}

						item.c = CONVERT_2_LUA;
						sprintf(item.inputName, "%s", strInput.c_str());
						item.inputSheetIdx = nSheetIdx;
						item.inputTitleRowId = nTitleRowId;
						sprintf(item.inputFields, "%s", strFields.c_str());
						sprintf(item.inputKeys, "%s", strKeys.c_str());
						sprintf(item.outputName, "%s", strOutput.c_str());
						sprintf(item.luaTableName, "%s", strLuaTable.c_str());

						// fields format: "A, C-Z,..."
						{
							char * chArr[8][32] = { 0 };
							int arrN2[8] = { 0 };
							int n = split2d(strFields.c_str(), strFields.length(), chArr, 8, ',', '-', arrN2);
							int i, j;
							int nFromColumn = -1;
							int nToColumn = -1;
							for (i = 0; i < n; ++i) {
								int nField1 = column_to_num(chArr[i][0]);
								int nField2 = column_to_num(chArr[i][1]);
								nField2 = nField2 > 0 ? nField2 : nField1;

								for (j = nField1; j <= nField2; ++j) {
									metaTable.arr[j].isField = true;

									if (-1 == nFromColumn) {
										nFromColumn = j;
									}

									//
									nToColumn = j;
								}
							}

							// fields : from - to
							metaTable.fromColumn = nFromColumn;
							metaTable.toColumn = nToColumn;

							if (nFromColumn >= 0) {
								metaTable.fromColumn = nFromColumn;
								metaTable.arr[nFromColumn].isFromColumn = true;

							}

							if (nToColumn >= 0) {
								metaTable.toColumn = nToColumn;
								metaTable.arr[nToColumn].isToColumn = true;
							}
						}

						// fields format: "A, C-Z,..."
						{
							char * chArr[8][32] = { 0 };
							int arrN2[8] = { 0 };
							int n = split2d(strKeys.c_str(), strKeys.length(), chArr, 8, ',', '-', arrN2);
							int i, j;
							for (i = 0; i < n; ++i) {
								int nKey1 = column_to_num(chArr[i][0]);
								int nKey2 = column_to_num(chArr[i][1]);
								nKey2 = nKey2 > 0 ? nKey2 : nKey1;

								for (j = nKey1; j <= nKey2; ++j) {
									if (metaTable.arr[j].isField)
										metaTable.arr[j].isKey = true;
								}
							}
						}

						//
						_config.push_back(item);

						// truncate output file 
						if (strlen(item.outputName) > 1) {
							FILE *fp_to_truncate = fopen(item.outputName, "wb+");
							fclose(fp_to_truncate);
						}
						else {
							fprintf(stderr, "output file <%s> truncate failed!!!\n", item.outputName);

							FILE *ferr = fopen("xlsxconverttool.error", "at+");
							fprintf(ferr, "output file <%s> truncate failed!!!\n", item.outputName);
							fclose(ferr);
							exit(-1);
						}
					}
				}

			}
		}

		//
		fclose(fp);
	}

	// free filebuf
	delete[] filebuf;
}

void
CXlsxConvertTool::Usage(char *progName) {
	fprintf(stderr, "\nusage: %s <Excel xls file> [Output filename]\n", progName);
	exit(EXIT_FAILURE);
}

char *
CXlsxConvertTool::W2c(const wchar_t *wstr, char *cstr, int clenMax) {
	if (wstr) {
		int wlen = wcslen(wstr);
		int nbytes = WideCharToMultiByte(
			CP_UTF8,	// specify the code page used to perform the conversion
			0,			// no special flags to handle unmapped characters
			wstr,		// wide character string to convert
			wlen,		// the number of wide characters in that string
			NULL,		// no output buffer given, we just want to know how long it needs to be
			0,
			NULL,		// no replacement character given
			NULL	    // we don't want to know if a character didn't make it through the translation
			);

		// make sure the buffer is big enough for this, making it larger if necessary
		if (nbytes > clenMax)
			nbytes = clenMax;

		WideCharToMultiByte(CP_UTF8, 0, wstr, wlen, cstr, nbytes, NULL, NULL);
		return cstr;
	}
	return NULL;
}

wchar_t *
CXlsxConvertTool::C2w(const char *cstr, wchar_t *wstr, int wlenMax) {
	if (cstr) {
		size_t clen = strlen(cstr);
		int nwords = MultiByteToWideChar(CP_UTF8, 0, (const char *)cstr, (int)clen, NULL, 0);
		if (nwords >= wlenMax)
			nwords = wlenMax - 1;

		MultiByteToWideChar(CP_UTF8, 0, (const char *)cstr, (int)clen, wstr, (int)nwords);
		wstr[nwords] = 0;
		return wstr;
	}
	return NULL;
}

bool
CXlsxConvertTool::StartsWith(const TCHAR *str, const TCHAR *pattern) {
	while (*str == TCHAR(' ') || *str == TCHAR('\t')) {
		++str;
	}

	while (*pattern && *str == *pattern) {
		++str;
		++pattern;
	}
	return *pattern == 0;
}

/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
*
*
*/
int
_tmain(int argc, TCHAR *argv[]) {
	CXlsxConvertTool app;

	app.SetupMetaTable();
	{
		std::vector<config_item_t>::iterator it = app._config.begin(), itEnd = app._config.end();
		while (it != itEnd) {
			config_item_t& item = (*it);
			if (CONVERT_2_SQL == item.c) {
				CConvert2Sql conv(&app);
				conv.Convert(&item);
			}
			else if (CONVERT_2_LUA == item.c) {
				CConvert2Lua conv(&app);
				conv.Convert(&item);
			}

			//
			++it;
		}
	}
	return EXIT_SUCCESS;
}