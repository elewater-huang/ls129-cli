#ifndef LS129_PROTOCOL_H
#define LS129_PROTOCOL_H
#include <stddef.h>
#include <stdint.h>
#ifdef __cplusplus
extern "C" {
#endif

typedef struct {
    uint16_t irradiance;   /* register 201: integer real-time value */
    uint16_t maximum;      /* register 202: integer maximum value */
    uint32_t energy;       /* registers 203-204: integer energy value */
} LS129Measurement;

uint16_t ls129_crc16(const uint8_t *data, size_t length);
int ls129_build_read_request(uint8_t station, uint8_t request[8]);
/* Returns 0 on success, negative on invalid data. Response is 13 bytes. */
int ls129_decode_measurement(uint8_t station, const uint8_t *response,
                             size_t length, LS129Measurement *measurement);
#ifdef __cplusplus
}
#endif
#endif
