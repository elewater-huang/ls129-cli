/* Verify x86 stdcall exports load correctly without physical RS485 hardware. */
#define WINVER 0x0601
#define _WIN32_WINNT 0x0601
#include <windows.h>
#include <stdio.h>
typedef LONG (__stdcall *PFN_OPEN)(const char*,LONG*);
typedef LONG (__stdcall *PFN_CLOSE)(LONG);
typedef LONG (__stdcall *PFN_READ)(LONG,double*,double*);
typedef LONG (__stdcall *PFN_VER)(void);
int main(void)
{
    HMODULE dll = LoadLibraryA(".\\LS129Core.dll");
    PFN_OPEN fn_open;
    PFN_CLOSE fn_close;
    PFN_READ fn_read;
    PFN_VER fn_version;
    LONG session = 123;
    double a = 0, b = 0;
    if (!dll) { fputs("FAIL: DLL cannot be loaded\\n", stderr); return 1; }
    fn_open = (PFN_OPEN)GetProcAddress(dll, "LS129_Open");
    fn_close = (PFN_CLOSE)GetProcAddress(dll, "LS129_Close");
    fn_read = (PFN_READ)GetProcAddress(dll, "LS129_Read");
    fn_version = (PFN_VER)GetProcAddress(dll, "LS129_APIVersion");
    if (!fn_open || !fn_close || !fn_read || !fn_version) {
        fputs("FAIL: missing exports\\n", stderr); return 2;
    }
    if (fn_version() != 3) return 3;
    if (fn_open("file-does-not-exist.ini", &session) != 2 || session != 0) return 4;
    if (fn_read(123, &a, &b) != 6) return 5;
    if (fn_close(123) != 6) return 6;
    FreeLibrary(dll);
    puts("DLL ABI TEST PASS (x86 exports, stdcall, no hardware)");
    return 0;
}
