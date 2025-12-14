#ifndef XLL_DEFINITIONS_H
#define XLL_DEFINITIONS_H

#include <windows.h>

// Excel 12 (2007+) definitions
// Ref: Microsoft Excel XLL SDK

typedef wchar_t XCHAR;
typedef int RW;
typedef int COL;

typedef struct xloper12 {
    union {
        double num;
        XCHAR *str;
        BOOL xbool;
        int err;
        struct {
            struct xloper12 *lparray;
            RW rows;
            COL columns;
        } array;
    } val;
    DWORD xltype;
} XLOPER12, *LPXLOPER12;

#define xltypeNum     0x0001
#define xltypeStr     0x0002
#define xltypeBool    0x0004
#define xltypeErr     0x0010
#define xltypeMulti   0x0040
#define xltypeMissing 0x0080

#define xlfRegister   149
#define xlfGetName    41
#define xlFree        (0x4000 | 13)

// Callback signature
typedef int (__stdcall *PEXCEL12)(int xlfn, int count, LPXLOPER12 operRes, LPXLOPER12 opers[]);

#endif // XLL_DEFINITIONS_H
