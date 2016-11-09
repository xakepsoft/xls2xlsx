#include <stdio.h>
#include <stdlib.h>
#include <stdio.h>
#include <libgen.h>
#include <string.h>
#include <unistd.h>
#include <ctype.h>
#include <sysexits.h>
#include "xlsxwriter.h"
#include "freexl.h"

static void const *freexl_handle = NULL;
static lxw_workbook  *workbook = NULL;

char* strlwr(char* s) {
    char* tmp = s;
    for (;*tmp;++tmp)*tmp = tolower((unsigned char) *tmp);
    return s;
}

void err_exit(int err_code) {
    if( freexl_handle ) freexl_close (freexl_handle);
    if( workbook ) workbook_close(workbook);
    exit( err_code );
}

int main( int argc, char *argv[] )
{
    unsigned int worksheet_index;
    const char *utf8_worksheet_name;
    int tmp, ret;
    unsigned int info;
    unsigned int max_worksheet;
    unsigned int rows;
    unsigned short columns;
    unsigned int row;
    unsigned short col;
    FreeXL_CellValue cell;
    lxw_worksheet *worksheet;
    lxw_workbook_options options = {.constant_memory = LXW_TRUE, .tmpdir = NULL};
    int inlineStr_mode = 0;

    char *infile; 
    char *outfile;

    while ((tmp=getopt(argc,argv,"xh"))!=-1) {
        switch(tmp) {
            case 'x':
                inlineStr_mode = 1;
                break;
            default:
                goto help;
        }
    }

    if(optind>=argc) {
        help:
        printf("Usage:\n      xls2xlsx [-x] xlsfile [xlsxfile]\n\n");
        printf("Spreadsheet data conversion from xls to xlsx format\n\n");
        printf("Options:\n");
        printf("       -x    Experimental feature! Uses a lot less system resources to generate xlsx.\n");
        printf("             Also this feature is very useful when dealing with realy big xlsx files because\n");
        printf("             files generated this way use \"inline strings\" instead of slower \"shared strings\".\n");
        printf("             CAUTION!  Apple Numbers doesn't support EXCEL \"inline strings\" standard yet.\n\n");
        err_exit( EX_USAGE );
    }

    infile = argv[optind];
    outfile = NULL;
    while(++optind < argc )
        if( argv[optind][0] != '-' ) {
            outfile = argv[optind];
            break;
        }

    if(!outfile) {
        outfile = (char*)malloc(sizeof(char) * strlen(infile) + 7 );
        strcpy( outfile , basename( infile ) );

        tmp = strlen( outfile );
        if( tmp>3 && outfile[tmp-1] == 's' && outfile[tmp-2] == 'l' && outfile[tmp-3] == 'x' && outfile[tmp-4] == '.')
            strcat( outfile , "x" );
        else
            strcat( outfile , ".xlsx" );
    }

    ret = freexl_open (infile, &freexl_handle);
    if (ret != FREEXL_OK) {
        fprintf (stderr, "Cannot open input file: %s\n", infile);
        err_exit( EX_NOINPUT );
    }

    ret = freexl_get_info (freexl_handle, FREEXL_BIFF_PASSWORD, &info);
    if (ret != FREEXL_OK) {
        fprintf (stderr, "GET-INFO [FREEXL_BIFF_PASSWORD] Error: %d\n", ret);
        err_exit( EX_DATAERR );
    }

    if (info == FREEXL_BIFF_OBFUSCATED) {
        fprintf (stderr, "Password protected xls file: (not accessible)\n");
        err_exit( EX_NOPERM );
    }

    ret = freexl_get_info (freexl_handle, FREEXL_BIFF_SHEET_COUNT, &max_worksheet);
    if (ret != FREEXL_OK) {
        fprintf (stderr, "GET-INFO [FREEXL_BIFF_SHEET_COUNT] Error: %d\n", ret);
        err_exit( EX_DATAERR );
    }

    if(inlineStr_mode)
        workbook  = workbook_new_opt( outfile, &options);
    else
        workbook  = workbook_new( outfile );

    if(!workbook) {
        fprintf (stderr, "Cannot create workbook: %s\n", outfile);
        err_exit( EX_CANTCREAT );
    }

    for (worksheet_index = 0; worksheet_index < max_worksheet; worksheet_index++) {
        ret = freexl_get_worksheet_name (freexl_handle, worksheet_index, &utf8_worksheet_name);
        if (ret != FREEXL_OK) {
            fprintf (stderr, "GET-WORKSHEET-NAME Error: %d\n", ret);
            err_exit( EX_DATAERR );
        }
        ret = freexl_select_active_worksheet (freexl_handle, worksheet_index);
        if (ret != FREEXL_OK) {
            fprintf (stderr, "SELECT-ACTIVE_WORKSHEET Error: %d\n", ret);
            err_exit( EX_DATAERR );
        }
        ret = freexl_worksheet_dimensions (freexl_handle, &rows, &columns);
        if (ret != FREEXL_OK) {
            fprintf (stderr, "WORKSHEET-DIMENSIONS Error: %d\n", ret);
            err_exit( EX_DATAERR );
        }
        worksheet = workbook_add_worksheet(workbook, utf8_worksheet_name);
        if(!worksheet) {
            fprintf (stderr, "Cannot create worksheet: %s\n", outfile);
            err_exit( EX_CANTCREAT );
        }
        for (row = 0; row < rows; row++) {
            for (col = 0; col < columns; col++) {
                ret = freexl_get_cell_value (freexl_handle, row, col, &cell);
                if (ret != FREEXL_OK) {
                    fprintf (stderr, "CELL-VALUE-ERROR (r=%u c=%u): %d\n", row, col, ret);
                    err_exit( EX_DATAERR );
                }
                switch(cell.type) {
                    case FREEXL_CELL_INT:
                        worksheet_write_number( worksheet , row, col, cell.value.int_value , NULL);
                        break;
                    case FREEXL_CELL_DOUBLE:
                        worksheet_write_number( worksheet , row, col, cell.value.double_value , NULL);
                        break;
                    case FREEXL_CELL_TEXT:
                    case FREEXL_CELL_SST_TEXT:
                    case FREEXL_CELL_DATE:
                    case FREEXL_CELL_DATETIME:
                    case FREEXL_CELL_TIME:
                        worksheet_write_string( worksheet , row, col, cell.value.text_value , NULL);
                        break;
                    case FREEXL_CELL_NULL:
                    default:
                        break;
                }
            }
        }
    }
    err_exit( EX_OK );
}

