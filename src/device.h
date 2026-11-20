#ifndef LS129_DEVICE_H
#define LS129_DEVICE_H
#ifndef WINVER
#define WINVER 0x0601
#endif
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0601
#endif
#include <windows.h>
#include "protocol.h"

/* Shared configuration, serial-port transport and probe operations.
 * No Python, CLR, or background service. Windows 7 x86 target.
 */
typedef struct {
    char port[32];
    DWORD baud;
    uint8_t station;
    DWORD timeout_ms;
    double poll_hz;
    double factor;
    int calibration_enabled;
} Config;

int parse_ulong(const char *text, unsigned long *value);
int parse_double(const char *text, double *value);
int read_config(const char *ini_path, Config *cfg);
int validate(const Config *cfg);
HANDLE open_com(const Config *cfg);
int get_measurement(HANDLE port, const Config *cfg, LS129Measurement *m);
int reset_recount(HANDLE port, const Config *cfg);
double corrected(const Config *cfg, uint16_t raw);

/* Function 03 register 320 (0 = none, 1 = 50Hz, 2 = 60Hz).
 * Setter uses function 06 and performs readback verification.
 */
int get_smoothing(HANDLE port, const Config *cfg, uint16_t *mode);
int set_smoothing(HANDLE port, const Config *cfg, uint16_t mode);

#endif
