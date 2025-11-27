//
//  main.cpp
//  xls-parser
//
//  Created by miyako on 2025/10/02.
//

#include "xls-parser.h"

static void usage(void)
{
    fprintf(stderr, "Usage:  xls-parser -r -i in -o out -\n\n");
    fprintf(stderr, "text extractor for xls documents\n\n");
    fprintf(stderr, " -%c path    : %s\n", 'i' , "document to parse");
    fprintf(stderr, " -%c path    : %s\n", 'o' , "text output (default=stdout)");
    fprintf(stderr, " %c          : %s\n", '-' , "use stdin for input");
    fprintf(stderr, " -%c         : %s\n", 'r' , "raw text output (default=json)");
    fprintf(stderr, " -%c charset : %s\n", 'c' , "charset (default=iso-8859-1)");
    exit(1);
}

extern OPTARG_T optarg;
extern int optind, opterr, optopt;

#ifdef _WIN32
OPTARG_T optarg = 0;
int opterr = 1;
int optind = 1;
int optopt = 0;
int getopt(int argc, OPTARG_T *argv, OPTARG_T opts) {

    static int sp = 1;
    register int c;
    register OPTARG_T cp;
    
    if(sp == 1)
        if(optind >= argc ||
             argv[optind][0] != '-' || argv[optind][1] == '\0')
            return(EOF);
        else if(wcscmp(argv[optind], L"--") == NULL) {
            optind++;
            return(EOF);
        }
    optopt = c = argv[optind][sp];
    if(c == ':' || (cp=wcschr(opts, c)) == NULL) {
        ERR(L": illegal option -- ", c);
        if(argv[optind][++sp] == '\0') {
            optind++;
            sp = 1;
        }
        return('?');
    }
    if(*++cp == ':') {
        if(argv[optind][sp+1] != '\0')
            optarg = &argv[optind++][sp+1];
        else if(++optind >= argc) {
            ERR(L": option requires an argument -- ", c);
            sp = 1;
            return('?');
        } else
            optarg = argv[optind++];
        sp = 1;
    } else {
        if(argv[optind][++sp] == '\0') {
            sp = 1;
            optind++;
        }
        optarg = NULL;
    }
    return(c);
}
#define ARGS (OPTARG_T)L"i:o:c:-rh"
#else
#define ARGS "i:o:c:-rh"
#endif

struct Row {
    std::vector<std::string> cells;
};

struct Sheet {
    std::string name;
    std::vector<Row> rows;
};

struct Document {
    std::string type;
    std::vector<Sheet> sheets;
};

static void document_to_json(Document& document, std::string& text, bool rawText) {
    
    if(rawText){
        text = "";
        for (const auto &sheet : document.sheets) {
            bool multiline = false;
            for (const auto &row : sheet.rows) {
                std::string _text;
                bool continued = false;
                for (const auto &cell : row.cells) {
                    if(continued) {
                        _text += ",";
                    }
                    _text += cell;
                    continued = true;
                }
                
                if(_text.length() != 0){
                    if(multiline) {
                        text += "\n";
                    }
                    text += _text;
                    multiline = true;
                }
            }
        }
    }else{
        Json::StreamWriterBuilder writer;
        writer["indentation"] = "";
        Json::Value documentNode(Json::objectValue);
        documentNode["type"] = document.type;
        documentNode["pages"] = Json::arrayValue;
        
        for (const auto &sheet : document.sheets) {
            Json::Value sheetNode(Json::objectValue);
            Json::Value sheetMetaNode(Json::objectValue);
            sheetMetaNode["name"] = sheet.name;
            sheetNode["meta"] = sheetMetaNode;
            sheetNode["paragraphs"] = Json::arrayValue;
            
            for (const auto &row : sheet.rows) {
                Json::Value paragraphNode(Json::objectValue);
                    paragraphNode["values"] = Json::arrayValue;
                for (const auto &cell : row.cells) {
                    paragraphNode["values"].append(cell);
                }
                
                std::string values = Json::writeString(writer, paragraphNode["values"]);
                paragraphNode["text"] = values;
                sheetNode["paragraphs"].append(paragraphNode);
            }
            documentNode["pages"].append(sheetNode);
        }
        text = Json::writeString(writer, documentNode);
    }
}

static std::string conv(const std::string& input, const std::string& charset) {
    
    std::string str = input;
    
    iconv_t cd = iconv_open("utf-8", charset.c_str());
    
    if (cd != (iconv_t)-1) {
        size_t inBytesLeft = input.size();
        size_t outBytesLeft = (inBytesLeft * 4)+1;
        std::vector<char> output(outBytesLeft);
        char *inBuf = const_cast<char*>(input.data());
        char *outBuf = output.data();
        if (iconv(cd, &inBuf, &inBytesLeft, &outBuf, &outBytesLeft) != (size_t)-1) {
            str = std::string(output.data(), output.size() - outBytesLeft);
        }
        iconv_close(cd);
    }
    
    return str;
}

using namespace xls;

int main(int argc, OPTARG_T argv[]) {
    
    const OPTARG_T input_path  = NULL;
    const OPTARG_T output_path = NULL;
    
    std::vector<uint8_t>xls_data(0);

    int ch;
    std::string text;
    bool rawText = false;
    std::string charset = "iso-8859-1";
        
    while ((ch = getopt(argc, argv, ARGS)) != -1){
        switch (ch){
            case 'i':
                input_path  = optarg;
                break;
            case 'o':
                output_path = optarg;
                break;
            case '-':
            {
                std::vector<uint8_t> buf(BUFLEN);
                size_t n;
                
                while ((n = fread(buf.data(), 1, buf.size(), stdin)) > 0) {
                    xls_data.insert(xls_data.end(), buf.begin(), buf.begin() + n);
                }
            }
                break;
            case 'r':
                rawText = true;
                break;
            case 'h':
            default:
                usage();
                break;
        }
    }
        
    if((!xls_data.size()) && (input_path != NULL)) {
        FILE *f = _fopen(input_path, _rb);
        if(f) {
            _fseek(f, 0, SEEK_END);
            size_t len = (size_t)_ftell(f);
            _fseek(f, 0, SEEK_SET);
            xls_data.resize(len);
            fread(xls_data.data(), 1, xls_data.size(), f);
            fclose(f);
        }
    }
    
    if(!xls_data.size()) {
        usage();
    }
           
    
    Document document;
    
    xls_error_t errorCode;
    xlsWorkBook *pWB = xls_open_buffer((const unsigned char *)xls_data.data(),
                                       (size_t)xls_data.size(),
                                       (const char *)charset.c_str(), &errorCode);
    if(pWB) {
        
        document.type = "xls";
        
        errorCode = xls_parseWorkBook(pWB);
        if(errorCode) {
            std::cerr << "fail:xls_parseWorkBook(" << errorCode << ")" << xls_getError(errorCode) << std::endl;
        }else{
            for (uint16_t i = 0; i < pWB->sheets.count; ++i) {
                xlsWorkSheet* pWS = xls_getWorkSheet(pWB, i);
                if (!pWS) {
                    std::cerr << "fail:xls_getWorkSheet" << std::endl;
                    continue;
                }
                
                errorCode = xls_parseWorkSheet(pWS);
                
                if(errorCode) {
                    std::cerr << "fail:xls_parseWorkSheet(" << errorCode << ")" << xls_getError(errorCode) << std::endl;
                }else{
                    
                    Sheet sheet;
                    sheet.name = pWB->sheets.sheet[i].name;
                    document.sheets.push_back(sheet);
                    
                    for (xls::DWORD row = 0; row <= pWS->rows.lastrow; ++row) {
                        
                        Row _row;
                        
                        for (xls::DWORD col = 0; col <= pWS->rows.lastcol; ++col) {
                            xlsCell* cell = xls_cell(pWS, row, col);
                            if (cell && cell->str) {
                                std::string str = cell->str;
                                _row.cells.push_back(conv(str, charset));
                                
//                                _row.cells.push_back(cell->str);
//                                char* str = codepage_decode(cell->str, strlen(cell->str), pWB);
//                                _row.cells.push_back(str);
                            }
                        }
                        sheet.rows.push_back(_row);
                    }
                    
                    document.sheets.push_back(sheet);
                }
                xls_close_WS(pWS);
            }
        }
        xls_close_WB(pWB);
    }else{
        std::cerr << "fail:xls_open_buffer(" << errorCode << ")" << xls_getError(errorCode) << std::endl;
    }
        
    document_to_json(document, text, rawText);
    
    if(!output_path) {
        std::cout << text << std::endl;
    }else{
        FILE *f = _fopen(output_path, _wb);
        if(f) {
            fwrite(text.c_str(), 1, text.length(), f);
            fclose(f);
        }
    }

    return 0;
}
