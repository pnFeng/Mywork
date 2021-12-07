from ICs import DolphinSH
import time
import sys
import visa
from bit_manipulate import bm


class DMM:
    def __init__(self, device):
        rm = visa.ResourceManager()
        try:
            self.dmm1 = rm.open_resource(device)
        except Exception:
            print('Open Multimeter Failed')
        model = self.dmm1.query("*IDN?")
        print(model)

    def meas(self):
        meas1 = float(self.dmm1.query('MEAS:VOLT:DC?'))
        return meas1


class SHVID:
    def __init__(self):
        self.dsh = DolphinSH()
        ret = self.dsh.connect()
        if ret is False:
            print("Failed to connect")
            sys.exit(-1)

        self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.0)
        self.dsh.set_i2c_bitrate(100)

    def enable(self):
        # Un-lock DIMM vendor region(0x40~0x7f)
        self.dsh.i2c_write_reg(0x37, 0x73)
        self.dsh.i2c_write_reg(0x38, 0x94)
        self.dsh.i2c_write_reg(0x39, 0x40)

        # Un-lock PMIC vendor region(0x80~0xff)
        self.dsh.i2c_write_reg(0x37, 0x79)
        self.dsh.i2c_write_reg(0x38, 0xbe)
        self.dsh.i2c_write_reg(0x39, 0x10)

        # Diable OVP
        self.dsh.i2c_write_reg(0x56, 0xF0)

        # enable
        ret, data = self.dsh.i2c_read_reg(0x2F)
        data = bm.set_bit(data, 2)
        self.dsh.i2c_write_reg(0x2F, data)
        self.dsh.i2c_write_reg(0x32, 0x80)

    def Trim_REF(self, binary):
        ret0, data = self.dsh.i2c_read_reg(rgt)
        data_2bit = bm.get_bit(data, 'l', [7, 6])
        data = data_2bit[0] * 2 ** 7 + data_2bit[1] * 2 ** 6 + binary
        self.dsh.i2c_write_reg(rgt, data)
        return data

    def ND_Finding(self, thrsh):
        self.dsh.i2c_write_reg(0x52, 0x02)
        self.dsh.i2c_write_reg(0x53, thrsh)
        time.sleep(1)
        self.dsh.i2c_write_reg(0x52, 0x81)
        time.sleep(1)
        ret, self.ndlo = self.dsh.i2c_read_reg(0x55)
        print
        'REG 0x55_lowside, VALUE {}'.format(int(self.ndlo))
        ret, self.ndhi = self.dsh.i2c_read_reg(0x54)
        print
        'REG 0x54_highside, VALUE {}'.format(int(self.ndhi))


rgt = 0xa9  #

m1 = DMM('USB0::0x2A8D::0x1401::MY57218437::0::INSTR')
a = SHVID()
a.enable()
list = [0x00, 0x09, 0x12, 0x1B, 0x24]
thrsh, thrshmode = [], []
meas1 = m1.meas()
ret, data = a.dsh.i2c_read_reg(rgt)
thrsh.append(hex(data))
thrshmode.append('default ' + str(m1.meas()))
for j in list:
    print('.........j........', hex(j))
    ret, data = a.dsh.i2c_read_reg(rgt)
    data = bm.get_bit(data, 's', range(5, -1, -1))
    data = int(data, 2)
    init_value = bm.binary2value(bm.orgncode2cpmcode(data, 5), 5)
    for i in range(2 ** 5 - init_value):
        print
        'CNT {}'.format(i)
        setting = init_value
        init_value += 1
        setting = bm.orgncode2cpmcode(bm.value2binary(setting, 5), 5)
        val = a.Trim_REF(setting)
        a.ND_Finding(j)
        m1.meas()
        if a.ndhi != 0:
            print
            '{} finding the high threshold {}'.format(hex(j), int(a.ndhi))
            thrshmode.append('hi ' + hex(j) + '...' + str(m1.meas()))
            thrsh.append(hex(val))
            print
            'Hi {}'.format(val)
            break
    # for i in range(init_value + 2 ** 5 + 1):
    #     print 'CNT {}'.format(i)
    #     setting = init_value
    #     init_value -= 1
    #     setting = bm.orgncode2cpmcode(bm.value2binary(setting, 5), 5)
    #     val = a.Trim_REF(setting)
    #     a.ND_Finding(j)
    #     m1.meas()
    #     if a.ndlo != 0:
    #         print '{} finding the low threshold {}'.format(hex(j), int(a.ndlo))
    #         thrshmode.append('lo ' + hex(j) + '...' + str(m1.meas()))
    #         thrsh.append(hex(val))
    #         print 'lo {}'.format(val)
    #         break
print
'thrsh {}'.format(thrsh)
print
'thrshmode {}'.format(thrshmode)
