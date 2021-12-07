import sys
import time
from ICs import DolphinSH
from bit_manipulate import bm
from Mylib.bench_resource import brsc

p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
el1 = brsc.Eload('USB0::0x0A69::0x083E::636002002010::0::INSTR')


# m1 = brsc.DMM('USB0::0x2A8D::0x1401::MY57215694::0::INSTR')
# m2 = brsc.DMM('USB0::0x2A8D::0x1401::MY57218437::0::INSTR')


class Piht:
    def __init__(self):
        self.dsh = DolphinSH()
        ret = self.dsh.connect()
        if ret is False:
            print("Failed to connect")
            sys.exit(-1)

        # self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.0)
        self.dsh.set_i2c_bitrate(100)

    def unlockDIMM(self, bEnable):
        if bEnable == 1:
            # # Un-lock DIMM vendor region(0x40~0x6f)
            self.dsh.i2c_write_reg(0x37, 0x73)
            self.dsh.i2c_write_reg(0x38, 0x94)
            self.dsh.i2c_write_reg(0x39, 0x40)

    def unlockIDT(self, bEnable):
        if bEnable == 1:
            # # Un-lock PMIC vendor region(0x70~0xff)
            self.dsh.i2c_write_reg(0x37, 0x79)
            self.dsh.i2c_write_reg(0x38, 0xbe)
            self.dsh.i2c_write_reg(0x39, 0x10)

    def securemode(self, bEnable):
        ret, data = self.dsh.i2c_read_reg(0x2F)
        if bEnable == 0:
            data = bm.set_bit(data, 2)
        elif bEnable == 1:
            data = bm.clear_bit(data, 2)
        self.dsh.i2c_write_reg(0x2F, data)

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)

    def GSI(self, bEnable):
        ret, data = self.dsh.i2c_read_reg(0x1B)
        assert ret == 0
        if bEnable == True:
            data = bm.set_bit(data, 3)
        elif bEnable == False:
            data = bm.clear_bit(data, 3)
        self.dsh.i2c_write_reg(0x1B, data)


p1.off(1)
time.sleep(1)
a = Piht()
p1.on(5, 4, 1)
time.sleep(1)
a.securemode(0)
a.unlockDIMM(1)
a.unlockIDT(0)
a.dsh.i2c_write_reg(0x16, 0x2b)  # mask current warning
# a.dsh.i2c_write_reg(0x18, 0x80)  # mask SWA CL
a.dsh.i2c_write_reg(0x4f, 0x01)  # dual phase
a.GSI(True)
a.enable(1)
time.sleep(1)

load = 7.5
while True:
    status = []
    load += 0.1
    el1.cc(load, 1)
    el1.on([1])
    time.sleep(0.5)
    for i in [0x08, 0x09, 0x0a, 0x0b, 0x33]:
        ret, data = a.dsh.i2c_read_reg(i)
        status.append(hex(data))
    print(status, load)
    el1.off([1])
