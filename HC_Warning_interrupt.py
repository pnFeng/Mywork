import time
import visa
from ICs import PineHurst
from bit_manipulate import bm
from Mylib.bench_resource import brsc


class PH:
    def __init__(self):
        self.dsh = PineHurst()
        ret = self.dsh.connect()
        assert ret == True

        self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.0)
        self.dsh.set_i2c_bitrate(100)

        # Un-lock DIMM vendor region(0x40~0x7f)
        self.dsh.i2c_write_reg(0x37, 0x73)
        self.dsh.i2c_write_reg(0x38, 0x94)
        self.dsh.i2c_write_reg(0x39, 0x40)

        # Un-lock PMIC vendor region(0x80~0xff)
        self.dsh.i2c_write_reg(0x37, 0x79)
        self.dsh.i2c_write_reg(0x38, 0xbe)
        self.dsh.i2c_write_reg(0x39, 0x10)

        ret, data = self.dsh.i2c_read_reg(0x2F)
        assert ret == 0
        data = bm.set_bit(data, 2)
        self.dsh.i2c_write_reg(0x2F, data)

    def __del__(self):
        ret = self.dsh.disconnect()
        print('Disconnect dongle')

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)


def polling_gsi(dmm, dongle):
    global data1
    print('read time {}'.format(int(time.time() * 1000)))
    for j in [0x08, 0x09, 0x0A, 0x0B, 0x33]:
        ret, data = dongle.dsh.i2c_read_reg(j)
        print('{} {}'.format(hex(j), hex(data)))
        data1 += data
    ret, data2 = dongle.dsh.i2c_read_reg(0x0C)
    ret, data3 = dongle.dsh.i2c_read_reg(0x0E)
    ret, data4 = dongle.dsh.i2c_read_reg(0x0F)
    print('cur_A {} cur_B {} cur_C {}'.format(data2, data3, data4))
    if data1 != 0:
        return 1
    else:
        return 0


def on_off(dongle):
    print('on_off time {}'.format(int(time.time() * 1000)))
    global data1
    i = 0
    while 1:
        i += 1
        dongle.dsh.i2c_write_reg(0x32, 0x80)
        print('-----------------on----------------------', i)
        time.sleep(0.1)
        dongle.dsh.i2c_write_reg(0x32, 0x00)
        print('-----------------off----------------------', i)
        time.sleep(0.1)
        if data1 != 0:
            return 1
        else:
            return 0


def main():
    m1 = brsc.DMM('USB0::0x2A8D::0x1401::MY57215694::0::INSTR')
    a = PH()
    a.dsh.i2c_write_reg(0x2F, 0x06)
    a.dsh.i2c_write_reg(0x20, 0xCF)
    # a.dsh.i2c_write_reg(0x2D, 0x2F)
    # a.dsh.i2c_write_reg(0x1C, 0xF0) # current level
    a.dsh.i2c_write_reg(0x30, 0x83)  # frequency
    # a.dsh.i2c_write_reg(0x92, 0x00)
    # a.dsh.i2c_write_reg(0xa0, 0x00)
    # a.dsh.i2c_write_reg(0xa8, 0x00)
    a.dsh.i2c_write_reg(0x14, 0x01)
    i, j = 0, 0
    while 1:
        i += 1
        start = time.time()
        print('-----------------on----------------------', i)
        a.enable(True)
        while 1:  # time.time() - start < 0.05:
            j += 1
            print(j)
            ret = polling_gsi(m1, a)
            if ret == 1:
                return
        start = time.time()
        a.enable(False)
        print('-----------------off----------------------', i)
        while time.time() - start < 0.05:
            polling_gsi(m1, a)
            if ret == 1:
                return


if __name__ == '__main__':
    data1 = 0
    main()
