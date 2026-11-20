/* LS129 core shared between command-line tool and VBA DLL. */
#define WINVER 0x0601
#define _WIN32_WINNT 0x0601
#define _CRT_SECURE_NO_WARNINGS
#include "device.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <float.h>
#include <errno.h>
static void read_ini_string(const char *ini, const char *section,
                            const char *key, const char *def, char *out, DWORD size)
{
    GetPrivateProfileStringA(section, key, def, out, size, ini);
}

int parse_ulong(const char *str, unsigned long *out)
{
    char *end = NULL;
    unsigned long v;
    if (str == NULL || *str == '\0' || *str == '-') return 0;
    errno = 0;
    v = strtoul(str, &end, 10);
    if (errno != 0 || end == str || *end != '\0') return 0;
    *out = v;
    return 1;
}

int parse_double(const char *str, double *out)
{
    char *end = NULL;
    double v;
    if (str == NULL || *str == '\0') return 0;
    errno = 0;
    v = strtod(str, &end);
    if (errno != 0 || end == str || *end != '\0' || !(_finite(v))) return 0;
    *out = v;
    return 1;
}

int read_config(const char *ini, Config *cfg)
{
    char tmp[64];
    unsigned long iv;
    int i;
    const char *section[] = {"Communication", "Communication", "Communication",
                             "Acquisition", "Calibration", "Calibration"};
    const char *key[] = {"baud_rate", "station_id", "timeout_ms", "poll_hz",
                         "factor", "enabled"};
    const char *def[] = {"9600", "1", "1000", "5", "1.0", "0"};
    read_ini_string(ini, "Communication", "port", "COM3", cfg->port, sizeof(cfg->port));
    for (i = 0; i < 6; ++i) {
        read_ini_string(ini, section[i], key[i], def[i], tmp, sizeof(tmp));
        if (i == 3 || i == 4) {
            double v;
            if (!parse_double(tmp, &v)) return 0;
            if (i == 3) cfg->poll_hz = v; else cfg->factor = v;
        } else {
            if (!parse_ulong(tmp, &iv)) return 0;
            if (i == 0) cfg->baud = (DWORD)iv;
            if (i == 1) { if (iv > 255) return 0; cfg->station = (uint8_t)iv; }
            if (i == 2) cfg->timeout_ms = (DWORD)iv;
            if (i == 5) cfg->calibration_enabled = (int)iv;
        }
    }
    return 1;
}

static int valid_port(const char *port)
{
    unsigned long number;
    if (strlen(port) < 4 || strlen(port) > 7) return 0;
    if ((port[0] != 'C' && port[0] != 'c') ||
        (port[1] != 'O' && port[1] != 'o') ||
        (port[2] != 'M' && port[2] != 'm')) return 0;
    if (!parse_ulong(port + 3, &number)) return 0;
    return number >= 1 && number <= 256;
}

int validate(const Config *cfg)
{
    if (!valid_port(cfg->port) || cfg->station < 1 || cfg->station > 247) return 0;
    if (cfg->baud != 4800 && cfg->baud != 9600 &&
        cfg->baud != 19200 && cfg->baud != 38400) return 0;
    if (cfg->timeout_ms < 100 || cfg->timeout_ms > 60000) return 0;
    if (cfg->poll_hz <= 0.0 || cfg->poll_hz > 32.0) return 0;
    if (cfg->factor <= 0.0 || cfg->factor > 1000000.0) return 0;
    if (cfg->calibration_enabled != 0 && cfg->calibration_enabled != 1) return 0;
    return 1;
}


HANDLE open_com(const Config *cfg)
{
    char name[48];
    HANDLE h;
    DCB dcb;
    COMMTIMEOUTS t;
    snprintf(name, sizeof(name), "\\\\.\\%s", cfg->port);
    h = CreateFileA(name, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (h == INVALID_HANDLE_VALUE) return h;
    memset(&dcb, 0, sizeof(dcb));
    dcb.DCBlength = sizeof(dcb);
    if (!GetCommState(h, &dcb)) { CloseHandle(h); return INVALID_HANDLE_VALUE; }
    dcb.BaudRate = cfg->baud;
    dcb.ByteSize = 8;
    dcb.Parity = NOPARITY;
    dcb.StopBits = ONESTOPBIT;
    dcb.fBinary = TRUE;
    dcb.fParity = FALSE;
    dcb.fOutxCtsFlow = FALSE;
    dcb.fOutxDsrFlow = FALSE;
    dcb.fDtrControl = DTR_CONTROL_DISABLE;
    dcb.fRtsControl = RTS_CONTROL_DISABLE; /* adapter must auto-switch RS485 direction */
    dcb.fOutX = FALSE;
    dcb.fInX = FALSE;
    if (!SetCommState(h, &dcb)) { CloseHandle(h); return INVALID_HANDLE_VALUE; }
    if (!SetupComm(h, 512, 512)) { /* non-fatal; some adapters ignore buffer sizes */ }
    memset(&t, 0, sizeof(t));
    t.ReadIntervalTimeout = 50;
    t.ReadTotalTimeoutMultiplier = 0;
    t.ReadTotalTimeoutConstant = 100;
    t.WriteTotalTimeoutMultiplier = 0;
    t.WriteTotalTimeoutConstant = 1000;
    if (!SetCommTimeouts(h, &t)) { CloseHandle(h); return INVALID_HANDLE_VALUE; }
    PurgeComm(h, PURGE_RXCLEAR | PURGE_TXCLEAR);
    return h;
}

static int read_exact(HANDLE h, uint8_t *dest, size_t expected, DWORD timeout_ms)
{
    DWORD start = GetTickCount();
    size_t total = 0;
    while (total < expected) {
        DWORD got = 0;
        DWORD slice = (DWORD)(expected - total);
        if ((DWORD)(GetTickCount() - start) > timeout_ms) return -1;
        if (!ReadFile(h, dest + total, slice, &got, NULL)) return -2;
        total += got;
        if (got == 0) Sleep(5);
    }
    return 0;
}

int get_measurement(HANDLE h, const Config *cfg, LS129Measurement *value)
{
    uint8_t req[8];
    uint8_t response[13];
    DWORD sent = 0;
    int result;
    uint16_t crc;
    if (ls129_build_read_request(cfg->station, req) != 0) return -10;
    PurgeComm(h, PURGE_RXCLEAR);
    if (!WriteFile(h, req, sizeof(req), &sent, NULL) || sent != sizeof(req)) return -11;
    if (!FlushFileBuffers(h)) return -12;
    result = read_exact(h, response, 3, cfg->timeout_ms);
    if (result) return -20;
    if (response[0] != cfg->station) return -21;
    if (response[1] == (uint8_t)(0x80u | 3u)) {
        result = read_exact(h, response + 3, 2, cfg->timeout_ms);
        if (result) return -22;
        crc = ls129_crc16(response, 3);
        if (response[3] != (uint8_t)(crc & 0xFFu) || response[4] != (uint8_t)(crc >> 8))
            return -23;
        return -100 - (int)response[2]; /* Modbus exception code */
    }
    if (response[1] != 3 || response[2] != 8) return -24;
    result = read_exact(h, response + 3, 10, cfg->timeout_ms);
    if (result) return -25;
    return ls129_decode_measurement(cfg->station, response, sizeof(response), value);
}

/* Recount on the configured station; require a valid function 06 echo. */
int reset_recount(HANDLE h, const Config *cfg)
{
    uint8_t request[8];
    uint8_t reply[8];
    DWORD sent = 0;
    int status;
    if (ls129_build_recount_request(cfg->station, request) != 0) return -10;
    if (!PurgeComm(h, PURGE_RXCLEAR)) return -11;
    if (!WriteFile(h, request, sizeof(request), &sent, NULL) ||
        sent != sizeof(request)) return -12;
    if (!FlushFileBuffers(h)) return -13;
    status = read_exact(h, reply, 2, cfg->timeout_ms);
    if (status) return -20;
    if (reply[0] != cfg->station) return -21;
    if (reply[1] == 0x86) {
        status = read_exact(h, reply + 2, 3, cfg->timeout_ms);
        if (status) return -22;
        return ls129_validate_recount_response(cfg->station, reply, 5);
    }
    if (reply[1] != 0x06) return -23;
    status = read_exact(h, reply + 2, 6, cfg->timeout_ms);
    if (status) return -24;
    return ls129_validate_recount_response(cfg->station, reply, sizeof(reply));
}

double corrected(const Config *cfg, uint16_t raw)
{
    return (double)raw * (cfg->calibration_enabled ? cfg->factor : 1.0);
}



/* One register read for verification of writable LS129 settings. */
static int read_register16(HANDLE h, const Config *cfg, uint16_t reg, uint16_t *out)
{
    uint8_t req[8] = {0}, reply[7] = {0};
    DWORD sent = 0;
    uint16_t crc;
    int rc;
    if (out == NULL || reg == 0) return -10;
    req[0] = cfg->station;
    req[1] = 3;
    req[2] = (uint8_t)(reg >> 8);
    req[3] = (uint8_t)reg;
    req[4] = 0;
    req[5] = 1;
    crc = ls129_crc16(req, 6);
    req[6] = (uint8_t)crc;
    req[7] = (uint8_t)(crc >> 8);
    if (!PurgeComm(h, PURGE_RXCLEAR)) return -11;
    if (!WriteFile(h, req, 8, &sent, NULL) || sent != 8) return -12;
    if (!FlushFileBuffers(h)) return -13;
    rc = read_exact(h, reply, 2, cfg->timeout_ms);
    if (rc) return -20;
    if (reply[0] != cfg->station) return -21;
    if (reply[1] == 0x83) {
        rc = read_exact(h, reply + 2, 3, cfg->timeout_ms);
        if (rc) return -22;
        crc = ls129_crc16(reply, 3);
        if (reply[3] != (uint8_t)crc || reply[4] != (uint8_t)(crc >> 8))
            return -23;
        return -100 - (int)reply[2];
    }
    if (reply[1] != 3) return -24;
    rc = read_exact(h, reply + 2, 5, cfg->timeout_ms);
    if (rc) return -25;
    if (reply[2] != 2) return -26;
    crc = ls129_crc16(reply, 5);
    if (reply[5] != (uint8_t)crc || reply[6] != (uint8_t)(crc >> 8))
        return -27;
    *out = (uint16_t)(((uint16_t)reply[3] << 8) | reply[4]);
    return 0;
}

int get_smoothing(HANDLE port, const Config *cfg, uint16_t *mode)
{
    return read_register16(port, cfg, 320, mode);
}

int set_smoothing(HANDLE port, const Config *cfg, uint16_t mode)
{
    uint8_t req[8] = {0};
    uint8_t reply[8] = {0};
    uint16_t crc, actual;
    DWORD sent = 0;
    int rc;
    if (mode > 2) return -10;
    req[0] = cfg->station;
    req[1] = 6;
    req[2] = 0x01;
    req[3] = 0x40;
    req[4] = 0;
    req[5] = (uint8_t)mode;
    crc = ls129_crc16(req, 6);
    req[6] = (uint8_t)crc;
    req[7] = (uint8_t)(crc >> 8);
    if (!PurgeComm(port, PURGE_RXCLEAR)) return -11;
    if (!WriteFile(port, req, 8, &sent, NULL) || sent != 8) return -12;
    if (!FlushFileBuffers(port)) return -13;
    rc = read_exact(port, reply, 2, cfg->timeout_ms);
    if (rc) return -20;
    if (reply[0] != cfg->station) return -21;
    if (reply[1] == 0x86) {
        rc = read_exact(port, reply + 2, 3, cfg->timeout_ms);
        if (rc) return -22;
        crc = ls129_crc16(reply, 3);
        if (reply[3] != (uint8_t)crc || reply[4] != (uint8_t)(crc >> 8))
            return -23;
        return -100 - (int)reply[2];
    }
    if (reply[1] != 6) return -24;
    rc = read_exact(port, reply + 2, 6, cfg->timeout_ms);
    if (rc) return -25;
    crc = ls129_crc16(reply, 6);
    if (reply[6] != (uint8_t)crc || reply[7] != (uint8_t)(crc >> 8))
        return -26;
    if (memcmp(reply, req, 8) != 0) return -27;
    /* Give the firmware time to apply the setting before readback. */
    Sleep(100);
    rc = get_smoothing(port, cfg, &actual);
    if (rc) return rc;
    return actual == mode ? 0 : -28;
}

