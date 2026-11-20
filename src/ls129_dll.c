/* LS129Core.dll: Win32 x86 stdcall ABI for Access 2000/2003 and Excel VBA.
 * This DLL shares device.c and protocol.c with LS129CLI.exe.
 * Not thread-safe: one device session per process; one owner per COM port.
 */
#define WINVER 0x0601
#define _WIN32_WINNT 0x0601
#define _CRT_SECURE_NO_WARNINGS
#include "device.h"
#include <string.h>

static HANDLE g_port = INVALID_HANDLE_VALUE;
static Config g_cfg;
static LONG g_session = 0;
static LONG g_next_session = 1;

/* Error policy: 0=success; 1=bad args; 2=INI missing/unreadable;
 * 3=invalid INI; 4=COM open failed; 5=already open;
 * 6=invalid session; negative codes come from Modbus/transport layer.
 */

LONG __stdcall LS129_Open(const char *ini_path, LONG *session)
{
    HANDLE port;
    Config cfg;
    DWORD attrs;
    if (session == NULL || ini_path == NULL || *ini_path == '\0') return 1;
    *session = 0;
    if (g_port != INVALID_HANDLE_VALUE) return 5;
    attrs = GetFileAttributesA(ini_path);
    if (attrs == INVALID_FILE_ATTRIBUTES || (attrs & FILE_ATTRIBUTE_DIRECTORY)) return 2;
    memset(&cfg, 0, sizeof(cfg));
    if (!read_config(ini_path, &cfg) || !validate(&cfg)) return 3;
    port = open_com(&cfg);
    if (port == INVALID_HANDLE_VALUE) return 4;
    g_cfg = cfg;
    g_port = port;
    if (g_next_session <= 0) g_next_session = 1;
    g_session = g_next_session++;
    *session = g_session;
    return 0;
}

static int valid_session(LONG session)
{
    return session != 0 && session == g_session && g_port != INVALID_HANDLE_VALUE;
}

LONG __stdcall LS129_Read(LONG session, double *raw, double *adjusted)
{
    LS129Measurement m;
    int rc;
    if (raw == NULL || adjusted == NULL) return 1;
    if (!valid_session(session)) return 6;
    rc = get_measurement(g_port, &g_cfg, &m);
    if (rc) return rc;
    *raw = (double)m.irradiance;
    *adjusted = corrected(&g_cfg, m.irradiance);
    return 0;
}

LONG __stdcall LS129_ReadAll(LONG session, double *raw, double *adjusted,
                              double *maximum, double *energy)
{
    LS129Measurement m;
    int rc;
    if (raw == NULL || adjusted == NULL || maximum == NULL || energy == NULL) return 1;
    if (!valid_session(session)) return 6;
    rc = get_measurement(g_port, &g_cfg, &m);
    if (rc) return rc;
    *raw = (double)m.irradiance;
    *adjusted = corrected(&g_cfg, m.irradiance);
    *maximum = (double)m.maximum;
    *energy = (double)m.energy;   /* preserves full unsigned 32-bit range */
    return 0;
}

LONG __stdcall LS129_Reset(LONG session)
{
    if (!valid_session(session)) return 6;
    return reset_recount(g_port, &g_cfg);
}

LONG __stdcall LS129_GetSmoothing(LONG session, LONG *mode)
{
    uint16_t value;
    int rc;
    if (mode == NULL) return 1;
    if (!valid_session(session)) return 6;
    rc = get_smoothing(g_port, &g_cfg, &value);
    if (rc) return rc;
    *mode = (LONG)value;
    return 0;
}

LONG __stdcall LS129_SetSmoothing(LONG session, LONG mode)
{
    if (!valid_session(session)) return 6;
    if (mode < 0 || mode > 2) return 1;
    return set_smoothing(g_port, &g_cfg, (uint16_t)mode);
}

LONG __stdcall LS129_Close(LONG session)
{
    if (!valid_session(session)) return 6;
    CloseHandle(g_port);
    g_port = INVALID_HANDLE_VALUE;
    g_session = 0;
    memset(&g_cfg, 0, sizeof(g_cfg));
    return 0;
}

LONG __stdcall LS129_APIVersion(void)
{
    return 3; /* SDK version 0.3 */
}
