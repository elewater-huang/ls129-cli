#include "../src/protocol.h"
#include <assert.h>
#include <stdio.h>
#include <string.h>

int main(void)
{
    uint8_t request[8] = {0};
    uint8_t expected[8] = {0x01, 0x03, 0x00, 0xC9, 0x00, 0x04, 0x94, 0x37};
    /* Real instrument example supplied during LS129 AccessPort testing. */
    uint8_t sample[13] = {0x01, 0x03, 0x08, 0x00, 0x15, 0x00, 0x1D,
                          0x00, 0x00, 0x0A, 0xAB, 0x7A, 0x0B};
    LS129Measurement value;
    assert(ls129_build_read_request(1, request) == 0);
    assert(memcmp(request, expected, sizeof(request)) == 0);
    assert(ls129_decode_measurement(1, sample, sizeof(sample), &value) == 0);
    assert(value.irradiance == 21);
    assert(value.maximum == 29);
    assert(value.energy == 2731u);
    /* Official function 06 recount request and its echoed reply. */
    {
        uint8_t recount[8] = {0};
        static const uint8_t expected_recount[8] =
            {0x01, 0x06, 0x00, 0x32, 0x00, 0x01, 0xE9, 0xC5};
        uint8_t exception_reply[5] = {0x01, 0x86, 0x02, 0, 0};
        uint16_t crc;
        assert(ls129_build_recount_request(1, recount) == 0);
        assert(memcmp(recount, expected_recount, 8) == 0);
        assert(ls129_validate_recount_response(1, recount, 8) == 0);
        assert(ls129_validate_recount_response(2, recount, 8) == -2);
        recount[4] = 1;
        assert(ls129_validate_recount_response(1, recount, 8) == -3);
        recount[4] = 0;
        recount[5] = 2;
        crc = ls129_crc16(recount, 6);
        recount[6] = (uint8_t)(crc & 0xFFu);
        recount[7] = (uint8_t)(crc >> 8);
        assert(ls129_validate_recount_response(1, recount, 8) == -4);
        crc = ls129_crc16(exception_reply, 3);
        exception_reply[3] = (uint8_t)(crc & 0xFFu);
        exception_reply[4] = (uint8_t)(crc >> 8);
        assert(ls129_validate_recount_response(1, exception_reply, 5) == -102);
    }
    sample[4] ^= 1;
    assert(ls129_decode_measurement(1, sample, sizeof(sample), &value) == -3);
    puts("Protocol tests PASS: read, recount, CRC and Modbus exception");
    return 0;
}
