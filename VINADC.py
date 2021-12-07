import time

import visa
from ICs import DolphinSH
from openpyxl import Workbook
from Mylib.bench_resource import brsc

from bit_manipulate import bm


class PHVINMeas:
    def __init__(self):
        self.dsh = DolphinSH()
        ret = self.dsh.connect()
        if ret is False:
            print("Failed to connect")
            sys.exit(-1)

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

        # VIN meas
        ret, data = self.dsh.i2c_read_reg(0x30)
        data = bm.set_bit(data, 7, 5, 3)
        self.dsh.i2c_write_reg(0x30, data)

    def enable(self):
        # enable
        ret, data = self.dsh.i2c_read_reg(0x2F)
        data = bm.set_bit(data, 2)
        self.dsh.i2c_write_reg(0x2F, data)
        self.dsh.i2c_write_reg(0x32, 0x80)

    def VINmeas(self):
        ret, data = self.dsh.i2c_read_reg(0x31)
        return data


ps = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
ps.on(5, 2, 1)
p = PHVINMeas()
p.enable()

VIN = 4.4
for i in range(11):
    VIN += 0.1
    ps.on(VIN, 2, 1)
    time.sleep(1)
    rslt = p.VINmeas()
    VIN_test = rslt * 0.07
    LSB = rslt - round(VIN / 0.07)
    print(VIN, hex(rslt), VIN_test, LSB)
