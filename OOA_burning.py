import sys
import time
from openpyxl import Workbook
# import dolphin
from ICs import DolphinSH
from bit_manipulate import bm
from Mylib.bench_resource import brsc
import prettytable as pt


class Piht:
    def __init__(self):
        self.dsh = DolphinSH()
        ret = self.dsh.connect()
        self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
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
        print(ret, data)
        if bEnable == 0:
            data = bm.bit_manipulate(0x2F, a, (2, 1))
        elif bEnable == 1:
            data = bm.bit_manipulate(0x2F, a, (2, 0))
        self.dsh.i2c_write_reg(0x2F, data)
        ret, data = self.dsh.i2c_read_reg(0x2F)
        print(ret, data)

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)


p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')

p1.on(5, 2, 2)
a = Piht()

# a.securemode(0)
a.unlockDIMM(1)
a.unlockIDT(1)
a.enable(0)
time.sleep(1)

ini = []
for i in range(256):
    ret, data = a.dsh.i2c_read_reg(i)
    ini.append(data)

bm.bit_manipulate(0x82, a, (6, 1), (5, 1), (4, 1))  # enable OOA and OOA Timer
bm.bit_manipulate(0x29, a, (5, 0), (4, 1))  # 1MHz
bm.bit_manipulate(0x4D, a, (5, 0), (4, 1))  # 1MHz
bm.bit_manipulate(0x2a, a, (5, 0), (4, 1))  # 1MHz
bm.bit_manipulate(0x4e, a, (5, 0), (4, 1))  # 1MHz

a.dsh.i2c_write_reg(0x39, 0x73)
time.sleep((0.5))
a.dsh.i2c_write_reg(0x39, 0x81)
time.sleep((0.5))
a.dsh.i2c_write_reg(0x39, 0x74)
time.sleep((0.5))
p1.off(2)
time.sleep(1)
p1.on(5, 2, 2)
time.sleep(1)

# a.securemode(0)
a.unlockDIMM(1)
a.unlockIDT(1)
a.enable(0)
time.sleep(1)

fin = []
for i in range(256):
    ret, data = a.dsh.i2c_read_reg(i)
    fin.append(data)

dif = {}
for i in range(256):
    if fin[i] != ini[i]:
        print('Reg {}, Ini {}, Fin {}'.format(hex(i), hex(ini[i]), hex(fin[i])))
