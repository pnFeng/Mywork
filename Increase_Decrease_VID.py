from Mylib.bench_resource import brsc
from ICs import PineHurst
from bit_manipulate import bm
import time


class Piht:
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

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)


path = './../' + 'EFF.xlsx'
p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
el1 = brsc.Eload('USB0::0x0A69::0x083E::636002002010::0::INSTR')
m1 = brsc.DMM('USB0::0x2A8D::0x1401::MY57215694::0::INSTR')
m2 = brsc.DMM('USB0::0x2A8D::0x1401::MY57218437::0::INSTR')
p1.on(5, 2, 1)

a = Piht()

# a.dsh.i2c_write_reg(0x27, 0x76)
ret, data = a.dsh.i2c_read_reg(0x27)
print('default', hex(data))

a.enable(1)
time.sleep(1)
for i in range(0x78, 0x00, -2):
    a.dsh.i2c_write_reg(0x27, i)
    time.sleep(0.1)
    ret, data = a.dsh.i2c_read_reg(0x27)
    print(hex(data))

for i in range(0x00, 0xff, 2):
    a.dsh.i2c_write_reg(0x27, i)
    time.sleep(0.1)
    ret, data = a.dsh.i2c_read_reg(0x27)
    print(hex(data))
