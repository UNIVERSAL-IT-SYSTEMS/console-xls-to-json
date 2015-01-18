#include "xlstojson.h"

#include <cstdio>
#include <sstream>
#include <vector>
#include <string>

using namespace std;

extern char *optarg;
extern int optind, opterr;

static void usage(void){
    fprintf(stderr, "usage: xlstojson [file]\n");
    exit(1);
}

#pragma mark -

int main(int argc, char *argv[])
{
    if(argc == 1) usage();
    
    int ch;
#ifdef _WIN32
    std::string path;    
#endif     
    while ((ch = getopt(argc, argv, "f:")) != -1){
                
        switch (ch){
                case 'f':
#ifdef _WIN32
                path = std::string(optarg);
                unescape_path(path);
                xlstojson((char *)path.c_str());  
#else
                    xlstojson(optarg);                
#endif                
                    break;
                                
            default:
                usage();
        }
    } 
    
	return 0;
}

#pragma mark -

void toWstr(const char *str, std::wstring &wstr){ 
    std::string temp(str);
    size_t size = temp.size() + sizeof(char);
    int wsize = size * 4;
    std::vector<char> buf(wsize);	            
    iconv_t icnv = iconv_open("WCHAR_T", "UTF-8");
    if(icnv != (iconv_t) -1){
        char *in = (char *)str;
        char *out =&buf[0];
        size_t insize = size -sizeof(char);
        size_t outsize = wsize -sizeof(wchar_t);                              
        size_t len = iconv(icnv, &in, &insize, &out, &outsize);
        iconv_close(icnv);        
        wstr = std::wstring((const wchar_t *)&buf[0], len);         
    }
}

void xlstojson(char *path){
    const void *h;
    int error = freexl_open(path, &h);
    if(FREEXL_OK == error){
        unsigned int worksheetCount;
        error = freexl_get_info(h, FREEXL_BIFF_SHEET_COUNT, &worksheetCount);
        if(FREEXL_OK == error){            
            JSONNODE *n = json_new(JSON_NODE);
            unsigned int worksheetEncoding;
            error = freexl_get_info(h, FREEXL_BIFF_CODEPAGE, &worksheetEncoding);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"encoding", worksheetEncoding));
            }
            unsigned int worksheetBiffVersion;
            error = freexl_get_info(h, FREEXL_BIFF_VERSION, &worksheetBiffVersion);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"version", worksheetBiffVersion));
            }
            unsigned int worksheetBiffMaxRecordSize;
            error = freexl_get_info(h, FREEXL_BIFF_MAX_RECSIZE, &worksheetBiffMaxRecordSize);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"record_max_size", worksheetBiffMaxRecordSize));
            }            
            unsigned int worksheetDateMode;
            error = freexl_get_info(h, FREEXL_BIFF_DATEMODE, &worksheetDateMode);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"date_mode", worksheetDateMode));
            }              
            unsigned int worksheetCfbfVersion;
            error = freexl_get_info(h, FREEXL_CFBF_VERSION, &worksheetCfbfVersion);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"cfbf_version", worksheetCfbfVersion));
            }            
            unsigned int worksheetCfbfSectorSize;
            error = freexl_get_info(h, FREEXL_CFBF_SECTOR_SIZE, &worksheetCfbfSectorSize);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"cfbf_sector_size", worksheetCfbfSectorSize));
            }              
            unsigned int worksheetCfbfFatCount;
            error = freexl_get_info(h, FREEXL_CFBF_FAT_COUNT, &worksheetCfbfFatCount);
            if(FREEXL_OK == error){ 
                json_push_back(n, json_new_i(L"cfbf_fat_count", worksheetCfbfFatCount));
            }             
             
            JSONNODE *sheets = json_new(JSON_ARRAY);
            json_set_name(sheets, L"Sheets");            
            json_push_back(n, sheets);
            
            for(unsigned int i = 0; i < worksheetCount;++i){
                const char *worsheetName;
                error = freexl_get_worksheet_name(h, i, &worsheetName);
                if(FREEXL_OK == error){
                    
                    std::wstring wworsheetName;
                    toWstr(worsheetName, wworsheetName);
                    JSONNODE *sheet = json_new(JSON_NODE);
                    JSONNODE *index = json_new_i(L"index", i);
                    JSONNODE *name = json_new_a(L"name", wworsheetName.c_str());
                    json_push_back(sheet, index);
                    json_push_back(sheet, name);
                    json_push_back(sheets, sheet);
                    
                    error = freexl_select_active_worksheet(h, i);
                    if(FREEXL_OK == error){
                    
                        unsigned int countRows;
                        unsigned short countColumns;
                        error = freexl_worksheet_dimensions(h, &countRows, &countColumns);
                        if(FREEXL_OK == error){
                            JSONNODE *rows = json_new(JSON_ARRAY);
                            json_set_name(rows, L"Rows"); 
                            json_push_back(sheet, rows); 
                            
                            for(unsigned int r = 0; r < countRows;++r){
                                
                                JSONNODE *row = json_new(JSON_NODE);
                                JSONNODE *index = json_new_i(L"index", r);
                                json_push_back(row, index);
                                json_push_back(rows, row);         
                                
                                JSONNODE *columns = json_new(JSON_ARRAY);
                                json_set_name(columns, L"Columns"); 
                                json_push_back(row, columns); 
                                
                                for(unsigned int c = 0; c < countColumns;++c){
                                    
                                    JSONNODE *cell = json_new(JSON_NODE);
                                    JSONNODE *index = json_new_i(L"index", c);
                                    json_push_back(cell, index);
                                    json_push_back(columns, cell);
                                    
                                    FreeXL_CellValue value;
                                    
                                    error = freexl_get_cell_value(h, r, c, &value);
                                    if(FREEXL_OK == error){
                                        std::wstring wvalue;
                                        JSONNODE *nvalue;
                                        switch (value.type){
                                            case FREEXL_CELL_INT:
                                                json_push_back(cell, json_new_i(L"value", value.value.int_value));
                                                break;
                                            case FREEXL_CELL_DOUBLE:
                                                json_push_back(cell, json_new_f(L"value", value.value.double_value));
                                                break;
                                            case FREEXL_CELL_TEXT:
                                            case FREEXL_CELL_SST_TEXT:
                                            case FREEXL_CELL_DATE:
                                            case FREEXL_CELL_DATETIME:
                                            case FREEXL_CELL_TIME: 
                                                toWstr(value.value.text_value, wvalue);
                                                json_push_back(cell, json_new_a(L"value", wvalue.c_str()));
                                                break;                    
                                            case FREEXL_CELL_NULL:
                                                nvalue = json_new(JSON_NULL);
                                                json_set_name(nvalue, L"value"); 
                                                json_nullify(nvalue);
                                                json_push_back(cell, nvalue);
                                                break;        
                                        }  
                                    } 
                                }
                            }    
                        }
                    }
                }
            }
            
            json_char *json = json_write(n);
            printf("%ls", json);
            json_free(json);
            json_delete(n);
        }
    }
}