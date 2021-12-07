from Mylib.bench_resource import brsc
from ICs import PineHurst
from bit_manipulate import bm
import time

status_rgt = [0x08, 0x09, 0x0a, 0x0b, 0x33]
clear_rgt = [0x10, 0x11, 0x12, 0x13, 0x14]
mask_rgt = [0x15, 0x16, 0x17, 0x18, 0x19]
name = 'None'

statusmap = ['R8[7]VIN_PG', 'R8[6]CTS', 'R8[5]SWA_PG', 'R8[4]RSV', 'R8[3]SWB_PG', 'R8[2]SWC_PG', 'R8[1]RSV',
             'R8[0]VIN_OV',
             'R9[7]VIN_HTW', 'R9[6]RSV', 'R9[5]LDO1.8PG', 'R9[4]RSV', 'R9[3]SWAHiCurWar', 'R9[2]RSV',
             'R9[1]SWBHiCurWar', 'R9[0]SWCHiCurWar',
             'RA[7]SWAOV', 'RA[6]RSV', 'RA[5]SWBOV', 'RA[4]SWCOV', 'RA[3]PECErr', 'RA[2]ParityErr', 'RA[1]GlobalStatus',
             'RA[0]RSV',
             'RB[7]SWACurLim', 'RB[6]RSV', 'RB[5]SWBCurLim', 'RB[4]SWCCurLim', 'RB[3]SWAUV', 'RB[2]RSV', 'RB[1]SWBUV',
             'RB[0]SWCUV',
             'R33[7]TM', 'R33[6]TM', 'R33[5]TM', 'R33[4]RSV', 'R33[3]VINUV', 'R33[2]LDO1.1PG', 'R33[1]RSV', 'R33[0]RSV']


class Piht:
    def __init__(self):
        self.dsh = PineHurst()
        ret = self.dsh.connect()
        assert ret == True

        self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.0)
        self.dsh.set_i2c_bitrate(100)

    def unlock(self):
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

    def globalclear(self):
        self.dsh.i2c_write_reg(0x14, 0x01)

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)

    def MaskCtrl(self, Maskctrl):
        ret, data = self.dsh.i2c_read_reg(0x2F)
        if Maskctrl == 0b00:
            data = bm.clear_bit(data, 0, 1)
        elif Maskctrl == 0b01:
            data = bm.clear_bit(data, 1)
            data = bm.set_bit(data, 0)
        elif Maskctrl == 0b10:
            data = bm.set_bit(data, 1)
            data = bm.clear_bit(data, 0)
        elif Maskctrl == 0b11:
            print('Please Check MastCtrl')
            raise
        self.dsh.i2c_write_reg(0x2F, data)


def statuscheck(dongle):
    rgt = []
    statusindex = ''
    for i in status_rgt:
        ret, data = dongle.dsh.i2c_read_reg(i)
        rgt.append(hex(data))
        index = status_rgt.index(i)
        str1 = str(bin(data))[2:10].zfill(8)
        for index1, s in enumerate(str1):
            if s == '1':
                statusindex = statusindex + ';' + str(statusmap[index1 + index * 8])

    return rgt, statusindex


a = Piht()
while 1:
    rgt, index = statuscheck(a)
    print(rgt)
    if rgt[0] == '0x40' or rgt[1] == '0x80':
        print('Critical Temperature Triggered')
        break
