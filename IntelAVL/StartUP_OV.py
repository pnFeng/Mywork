import time
from ICs import PineHurst
from openpyxl import Workbook
from bit_manipulate import bm
from Mylib.bench_resource import brsc


class PHCurMeas:
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

    def GSI(self, bEnable):
        ret, data = self.dsh.i2c_read_reg(0x1B)
        assert ret == 0
        if bEnable == True:
            data = bm.set_bit(data, 3)
        elif bEnable == False:
            data = bm.clear_bit(data, 3)
        self.dsh.i2c_write_reg(0x1B, data)

    def enable(self, bEnable):
        ret, data = self.dsh.i2c_read_reg(0x32)
        if bEnable:
            data = bm.set_bit(data, 7)
            self.dsh.i2c_write_reg(0x32, data)
        else:
            data = bm.clear_bit(data, 7)
            self.dsh.i2c_write_reg(0x32, data)


el1 = brsc.Eload('USB0::0x0A69::0x083E::636002002010::0::INSTR')
p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
msk = 0

p1.on(5.5, 4, 1)
p = PHCurMeas()
p.GSI(True)
p.enable(True)
time.sleep(1)
p.enable(False)


def rgtset(dongle, rgt, set, *n):
    ret, data = dongle.dsh.i2c_read_reg(rgt)
    if set == 1:
        data = bm.set_bit(data, *n)
    elif set == 0:
        data = bm.clear_bit(data, *n)
    dongle.dsh.i2c_write_reg(rgt, data)


def SWA_OV(dongle, bMask):
    if bMask == 0:
        rgtset(dongle, 0x17, 0, 7)
    elif bMask == 1:
        rgtset(dongle, 0x17, 1, 7)
    rgtset(dongle, 0x15, 1, 5)
    rgtset(dongle, 0x82, 0, 1)  # disable Soft_OV
    p1.on(1.21, 2, 3)
    time.sleep(1)
    dongle.enable(True)
    time.sleep(1)


def SWB_OV(dongle, bMask):
    if bMask == 0:
        rgtset(dongle, 0x17, 0, 5)
    elif bMask == 1:
        rgtset(dongle, 0x17, 1, 5)
    rgtset(dongle, 0x15, 1, 3)
    rgtset(dongle, 0x82, 0, 1)  # disable Soft_OV
    p1.on(1.21, 2, 3)
    time.sleep(1)
    dongle.enable(True)
    time.sleep(1)


def SWC_OV(dongle, bMask):
    if bMask == 0:
        rgtset(dongle, 0x17, 0, 4)
    elif bMask == 1:
        rgtset(dongle, 0x17, 1, 4)
    rgtset(dongle, 0x15, 1, 2)
    rgtset(dongle, 0x82, 0, 1)  # disable Soft_OV
    p1.on(2, 2, 3)
    time.sleep(1)
    dongle.enable(True)
    time.sleep(1)


SWC_OV(p, msk)
