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
    sample[4] ^= 1;
    assert(ls129_decode_measurement(1, sample, sizeof(sample), &value) == -3);
    puts("Protocol tests PASS: request CRC, response CRC, measurement decode");
    return 0;
}
