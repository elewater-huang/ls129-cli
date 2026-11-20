/* LS129CLI V0.3 - Win32 native COM port reader, no Python runtime.
 * Windows 7 target, x86 build. Record auto-recounts energy and peak.
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
#include "device.h"

#define MAX_COUNT 10000000UL

static void usage(void)
{
    puts("LS129CLI V0.2 (record auto-recounts; Modbus RTU)");
    puts("  LS129CLI.exe read");
    puts("  LS129CLI.exe status");
    puts("  LS129CLI.exe reset (recount energy and maximum)");
    puts("  LS129CLI.exe record --count 100 --out data.csv");
    puts("  LS129CLI.exe read --port COM4 --station 1 --baud 9600");
    puts("  LS129CLI.exe read --config C:\\path\\LS129.ini");
    puts("  LS129CLI.exe selftest (no probe required)");
    puts("Options: --port, --station, --baud, --hz, --count, --out, --config");
    puts("Reads integers 201..204; record resets statistics via register 50.");
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

static int do_selftest(void)
{
    uint8_t req[8];
    static const uint8_t expected[8] = {1, 3, 0, 0xC9, 0, 4, 0x94, 0x37};
    static const uint8_t reply[13] = {1, 3, 8, 0, 0x15, 0, 0x1D,
                                      0, 0, 0x0A, 0xAB, 0x7A, 0x0B};
    LS129Measurement m;
    static const uint8_t expected_reset[8] = {1, 6, 0, 0x32, 0, 1, 0xE9, 0xC5};
    uint8_t reset_req[8];
    if (ls129_build_recount_request(1, reset_req) != 0 ||
        memcmp(reset_req, expected_reset, 8) != 0 ||
        ls129_validate_recount_response(1, expected_reset, 8) != 0 ||
        ls129_build_read_request(1, req) != 0 || memcmp(req, expected, 8) ||
        ls129_decode_measurement(1, reply, 13, &m) != 0 ||
        m.irradiance != 21 || m.maximum != 29 || m.energy != 2731u) {
        fputs("SELFTEST FAIL\n", stderr);
        return 1;
    }
    puts("SELFTEST PASS (CRC and decoding; no hardware tested)");
    return 0;
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
    if (strcmp(mode, "read") && strcmp(mode, "status") &&
        strcmp(mode, "record") && strcmp(mode, "reset")) {
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
    if (!strcmp(mode, "reset")) {
        status = reset_recount(h, &cfg);
        CloseHandle(h);
        if (status) {
            fprintf(stderr, "LS129 recount failed (%d); reset not acknowledged.\n", status);
            return 6;
        }
        puts("Recount acknowledged: energy and maximum restarted on LS129.");
        return 0;
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
    /* Preflight output without truncating an earlier CSV.
     * An unsuccessful recount must never overwrite previous results.
     */
    fp = fopen(output_file, "ab");
    if (fp == NULL) {
        fprintf(stderr, "Cannot create/open output file: %s\n", output_file);
        CloseHandle(h); return 5;
    }
    fclose(fp);
    status = reset_recount(h, &cfg);
    if (status) {
        fprintf(stderr, "Record ABORTED: recount failed (%d); old CSV preserved.\n", status);
        CloseHandle(h); return 6;
    }
    fp = fopen(output_file, "wb");
    if (fp == NULL) {
        fprintf(stderr, "Cannot create output after recount: %s\n", output_file);
        CloseHandle(h); return 5;
    }
    puts("Recount acknowledged: new energy/maximum measurement period started.");
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
