/* LS129CLI V0.1 - Win32 native COM port reader, no Python runtime.
 * Windows 7 target, x86 build. Read-only by design for first release.
 */
#define WINVER 0x0601
#define _WIN32_WINNT 0x0601
#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <float.h>
#include <errno.h>
#include "protocol.h"

#define MAX_COUNT 10000000UL

typedef struct {
    char port[32];
    DWORD baud;
    uint8_t station;
    DWORD timeout_ms;
    double poll_hz;
    double factor;
    int calibration_enabled;
} Config;

static void usage(void)
{
    puts("LS129CLI V0.1 (read-only; Modbus RTU)");
    puts("  LS129CLI.exe read");
    puts("  LS129CLI.exe status");
    puts("  LS129CLI.exe record --count 100 --out data.csv");
    puts("  LS129CLI.exe read --port COM4 --station 1 --baud 9600");
    puts("  LS129CLI.exe read --config C:\\path\\LS129.ini");
    puts("  LS129CLI.exe selftest (no probe required)");
    puts("Options: --port, --station, --baud, --hz, --count, --out, --config");
    puts("Reads raw integer values from LS129 registers 201..204.");
}

static int find_default_ini(char *path, DWORD capacity)
{
    DWORD n = GetModuleFileNameA(NULL, path, capacity);
    char *last;
    if (!n || n >= capacity) return 0;
    last = strrchr(path, '\\');
    if (last == NULL) return 0;
    last[1] = '\0';
    if (strlen(path) + strlen("LS129.ini") + 1 > capacity) return 0;
    strcat(path, "LS129.ini");
    return 1;
}

static void read_ini_string(const char *ini, const char *section,
                            const char *key, const char *def, char *out, DWORD size)
{
    GetPrivateProfileStringA(section, key, def, out, size, ini);
}

static int parse_ulong(const char *str, unsigned long *out)
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

static int parse_double(const char *str, double *out)
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

static int read_config(const char *ini, Config *cfg)
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

static int validate(const Config *cfg)
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

static int do_selftest(void)
{
    uint8_t req[8];
    static const uint8_t expected[8] = {1, 3, 0, 0xC9, 0, 4, 0x94, 0x37};
    static const uint8_t reply[13] = {1, 3, 8, 0, 0x15, 0, 0x1D,
                                      0, 0, 0x0A, 0xAB, 0x7A, 0x0B};
    LS129Measurement m;
    if (ls129_build_read_request(1, req) != 0 || memcmp(req, expected, 8) ||
        ls129_decode_measurement(1, reply, 13, &m) != 0 ||
        m.irradiance != 21 || m.maximum != 29 || m.energy != 2731u) {
        fputs("SELFTEST FAIL\n", stderr);
        return 1;
    }
    puts("SELFTEST PASS (CRC and decoding; no hardware tested)");
    return 0;
}

static HANDLE open_com(const Config *cfg)
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

static int get_measurement(HANDLE h, const Config *cfg, LS129Measurement *value)
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

static double corrected(const Config *cfg, uint16_t raw)
{
    return (double)raw * (cfg->calibration_enabled ? cfg->factor : 1.0);
}

static void now_iso(char *buffer, size_t capacity)
{
    SYSTEMTIME t;
    GetLocalTime(&t);
    snprintf(buffer, capacity, "%04u-%02u-%02uT%02u:%02u:%02u.%03u",
             t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wSecond, t.wMilliseconds);
}

static void sleep_interval(double milliseconds)
{
    /* Handles arbitrarily slow positive poll_hz; Sleep() accepts DWORD only. */
    while (milliseconds >= 60000.0) { Sleep(60000); milliseconds -= 60000.0; }
    if (milliseconds > 0.5) Sleep((DWORD)(milliseconds + 0.5));
}

int main(int argc, char **argv)
{
    Config cfg;
    char ini[MAX_PATH];
    const char *mode;
    const char *output_file = "data.csv";
    unsigned long count = 100;
    int i;
    HANDLE h;
    LS129Measurement m;
    FILE *fp;
    int status;

    if (argc < 2 || !strcmp(argv[1], "--help") || !strcmp(argv[1], "help")) {
        usage(); return argc < 2 ? 1 : 0;
    }
    mode = argv[1];
    if (!strcmp(mode, "selftest")) return do_selftest();
    if (strcmp(mode, "read") && strcmp(mode, "status") && strcmp(mode, "record")) {
        usage(); return 2;
    }
    if (!find_default_ini(ini, sizeof(ini))) { fputs("Cannot locate application folder\n", stderr); return 2; }
    for (i = 2; i + 1 < argc; i += 2) {
        if (!strcmp(argv[i], "--config")) {
            if (strlen(argv[i+1]) >= sizeof(ini)) { fputs("Config path too long\n", stderr); return 2; }
            strcpy(ini, argv[i+1]);
        }
    }
    if (!read_config(ini, &cfg)) { fputs("Invalid LS129.ini value\n", stderr); return 2; }
    for (i = 2; i < argc; i += 2) {
        unsigned long v;
        if (i + 1 >= argc) { fputs("Option without value\n", stderr); return 2; }
        if (!strcmp(argv[i], "--config")) continue;
        if (!strcmp(argv[i], "--port")) {
            if (strlen(argv[i+1]) >= sizeof(cfg.port)) return 2;
            strcpy(cfg.port, argv[i+1]);
        } else if (!strcmp(argv[i], "--station")) {
            if (!parse_ulong(argv[i+1], &v) || v > 255) return 2;
            cfg.station = (uint8_t)v;
        } else if (!strcmp(argv[i], "--baud")) {
            if (!parse_ulong(argv[i+1], &v)) return 2;
            cfg.baud = (DWORD)v;
        } else if (!strcmp(argv[i], "--hz")) {
            if (!parse_double(argv[i+1], &cfg.poll_hz)) return 2;
        } else if (!strcmp(argv[i], "--count")) {
            if (!parse_ulong(argv[i+1], &count) || count < 1 || count > MAX_COUNT) return 2;
        } else if (!strcmp(argv[i], "--out")) {
            output_file = argv[i+1];
        } else {
            fprintf(stderr, "Unknown option: %s\n", argv[i]); return 2;
        }
    }
    if (!validate(&cfg)) {
        fputs("Invalid configuration. Expected COM1..COM256, baud 4800/9600/19200/38400,"
              " station 1..247, poll_hz >0 <=32, timeout 100..60000ms\n", stderr);
        return 2;
    }
    h = open_com(&cfg);
    if (h == INVALID_HANDLE_VALUE) {
        fprintf(stderr, "Cannot open/configure %s (Win32 error %lu); check USB-RS485 driver and port use\n",
                cfg.port, (unsigned long)GetLastError());
        return 3;
    }
    if (!strcmp(mode, "read") || !strcmp(mode, "status")) {
        status = get_measurement(h, &cfg, &m);
        CloseHandle(h);
        if (status) {
            fprintf(stderr, "LS129 read failed (%d). Check station, baud, wiring and CRC.\n", status);
            return 4;
        }
        if (!strcmp(mode, "read")) printf("%u\n", (unsigned)m.irradiance);
        else {
            printf("port=%s\nstation=%u\nbaud=%lu\n", cfg.port, (unsigned)cfg.station,
                   (unsigned long)cfg.baud);
            printf("raw_irradiance=%u\nraw_maximum=%u\nraw_energy=%lu\n",
                   (unsigned)m.irradiance, (unsigned)m.maximum, (unsigned long)m.energy);
            printf("corrected_irradiance=%.3f\n", corrected(&cfg, m.irradiance));
            puts("note=integer Modbus values; check probe model for engineering units");
        }
        return 0;
    }
    fp = fopen(output_file, "wb");
    if (fp == NULL) {
        fprintf(stderr, "Cannot create file: %s\n", output_file);
        CloseHandle(h); return 5;
    }
    fputs("timestamp_local,station,raw_irradiance,corrected_irradiance,raw_maximum,raw_energy\r\n", fp);
    for (i = 0; (unsigned long)i < count; ++i) {
        DWORD started = GetTickCount();
        char timestamp[48];
        status = get_measurement(h, &cfg, &m);
        if (status) {
            fprintf(stderr, "Record failed at sample %d (%d). File contains completed rows only.\n", i+1, status);
            fclose(fp); CloseHandle(h); return 4;
        }
        now_iso(timestamp, sizeof(timestamp));
        fprintf(fp, "%s,%u,%u,%.3f,%u,%lu\r\n", timestamp, (unsigned)cfg.station,
                (unsigned)m.irradiance, corrected(&cfg, m.irradiance),
                (unsigned)m.maximum, (unsigned long)m.energy);
        fflush(fp);
        if ((unsigned long)(i+1) < count) {
            double remain = 1000.0 / cfg.poll_hz - (double)(DWORD)(GetTickCount() - started);
            if (remain > 0.0) sleep_interval(remain);
        }
    }
    fclose(fp);
    CloseHandle(h);
    printf("Recorded %lu polls to %s (poll_hz=%.6g; repeated readings may occur)\n",
           count, output_file, cfg.poll_hz);
    return 0;
}
