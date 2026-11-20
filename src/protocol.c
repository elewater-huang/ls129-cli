#include "protocol.h"

uint16_t ls129_crc16(const uint8_t *data, size_t length)
{
    uint16_t crc = 0xFFFFu;
    size_t i;
    unsigned int bit;
    if (data == NULL) return 0;
    for (i = 0; i < length; ++i) {
        crc ^= data[i];
        for (bit = 0; bit < 8; ++bit) {
            if (crc & 1u) crc = (uint16_t)((crc >> 1) ^ 0xA001u);
            else crc >>= 1;
        }
    }
    return crc;
}

int ls129_build_read_request(uint8_t station, uint8_t request[8])
{
    uint16_t crc;
    if (station < 1 || station > 247 || request == NULL) return -1;
    /* Function 03, start address 201 (0x00C9), 4 registers. */
    request[0] = station;
    request[1] = 3;
    request[2] = 0;
    request[3] = 0xC9;
    request[4] = 0;
    request[5] = 4;
    crc = ls129_crc16(request, 6);
    request[6] = (uint8_t)(crc & 0xFFu);
    request[7] = (uint8_t)(crc >> 8);
    return 0;
}

int ls129_decode_measurement(uint8_t station, const uint8_t *response,
                             size_t length, LS129Measurement *measurement)
{
    uint16_t crc;
    if (response == NULL || measurement == NULL || length != 13) return -1;
    if (response[0] != station || response[1] != 3 || response[2] != 8) return -2;
    crc = ls129_crc16(response, length - 2);
    if (response[11] != (uint8_t)(crc & 0xFFu) ||
        response[12] != (uint8_t)(crc >> 8)) return -3;
    measurement->irradiance = (uint16_t)((uint16_t)response[3] << 8 | response[4]);
    measurement->maximum = (uint16_t)((uint16_t)response[5] << 8 | response[6]);
    measurement->energy = ((uint32_t)response[7] << 24) |
                          ((uint32_t)response[8] << 16) |
                          ((uint32_t)response[9] << 8) |
                          (uint32_t)response[10];
    return 0;
}
