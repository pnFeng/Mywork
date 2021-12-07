import time
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from Mylib.bench_resource import brsc
from ICs import PineHurst
from bit_manipulate import bm
from openpyxl.styles import colors, Font


class Piht:
    def __init__(self):
        self.dsh = PineHurst()
        ret = self.dsh.connect()
        assert ret == True

        self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.0)
        self.dsh.set_i2c_bitrate(1000)

    def unlock(self):
        # Un-lock DIMM vendor region(0x40~0x6f)
        self.dsh.i2c_write_reg(0x37, 0x73)
        self.dsh.i2c_write_reg(0x38, 0x94)
        self.dsh.i2c_write_reg(0x39, 0x40)

        # Un-lock PMIC vendor region(0x70~0xff)
        self.dsh.i2c_write_reg(0x37, 0x79)
        self.dsh.i2c_write_reg(0x38, 0xbe)
        self.dsh.i2c_write_reg(0x39, 0x10)

        ret, data = self.dsh.i2c_read_reg(0x2F)
        assert ret == 0
        data = bm.set_bit(data, 2)
        self.dsh.i2c_write_reg(0x2F, data)

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)


l = ['0xfe', '0xcc', '0xf0', '0xa', '0xee']

l1 = ['0x0', '0x0', '0x80', '0x0', '0x0']


def burning_MTP(vin, interval):
    datal = []
    for i in range(0x6a, 0x6F):
        _, data = a.dsh.i2c_read_reg(i)
        datal.append(hex(data))

    print('ini_reg_value {}'.format(datal))

    for i in range(0x60, 0x70):
        a.dsh.i2c_write_reg(i, 0xff)

    a.dsh.i2c_write_reg(0x39, 0x85)
    time.sleep(interval)
    p1.off(1)
    time.sleep(1)
    p1.on(vin, 2, 1)
    time.sleep(1)

    a.unlock()
    datal1 = []
    for i in range(0x6a, 0x6F):
        _, data = a.dsh.i2c_read_reg(i)
        datal1.append(hex(data))

    print('burning_power_cycle {}'.format(datal1))

    print('burning success {}'.format(l == datal1))


def erase(vin):
    a.unlock()
    j = 0
    for i in range(0x6a, 0x6F):
        a.dsh.i2c_write_reg(i, int(l1[j][2:4], 16))
        j += 1

    a.dsh.i2c_write_reg(0x39, 0x85)
    time.sleep(1)
    p1.off(1)
    time.sleep(1)
    p1.on(vin, 2, 1)
    time.sleep(1)

    a.unlock()
    datal1 = []
    for i in range(0x6a, 0x6F):
        _, data = a.dsh.i2c_read_reg(i)
        datal1.append(hex(data))

    print('erase power cycle {}'.format(datal1))
    print('erase success {}'.format(l1 == datal1))


def burning_i2c_disturb():
    datal = []
    for i in range(0x6a, 0x6F):
        _, data = a.dsh.i2c_read_reg(i)
        datal.append(hex(data))

    print('ini_reg_value {}'.format(datal))

    for i in range(0x60, 0x70):
        a.dsh.i2c_write_reg(i, 0xff)

    time.sleep(1)
    a.dsh.i2c_write_reg(0x39, 0x85)
    for i in range(0x40):
        # _, data = a.dsh.i2c_read_reg(i)
        a.dsh.i2c_write_reg(i, 0xff)


def power_cycle(p):
    p.off(1)
    time.sleep(0.5)
    i = 0
    while 1:
        i += 1
        print(i)
        p.on(5, 2, 1)
        time.sleep(0.2)
        p.off(1)
        time.sleep(0.5)


p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
# el1 = brsc.Eload('USB0::0x0A69::0x083E::636002002010::0::INSTR')
# m1 = brsc.DMM('USB0::0x2A8D::0x1401::MY57215694::0::INSTR')
# m2 = brsc.DMM('USB0::0x2A8D::0x1401::MY57218437::0::INSTR')

vin = 5
p1.on(vin, 2, 1)

# a = Piht()
# a.unlock()

# burning_MTP(vin, 1)
# time.sleep(1)
# p1.off(1)
# time.sleep(1)
# p1.on(5, 2, 1)
# time.sleep(1)
# erase(vin)

# burning_i2c_disturb()

power_cycle(p1)
