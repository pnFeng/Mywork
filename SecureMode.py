# -*- coding: utf-8 -*-

import sys
import time
import xlsxwriter
# import dolphin
from ICs import DolphinSH
from bit_manipulate import bm
from Mylib.bench_resource import brsc


class SHSecureMode:
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
        ret, data = self.dsh.i2c_read_reg(0x2F)
        print(hex(data))

    def enable(self, bEnable):
        if bEnable:
            self.dsh.i2c_write_reg(0x32, 0x80)
        else:
            self.dsh.i2c_write_reg(0x32, 0x00)


workbook = xlsxwriter.Workbook('../SecureMode.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()

cell_format1 = workbook.add_format({'bold': True,
                                    'font_color': 'black',
                                    'align': 'center',
                                    'valign': 'vcenter'
                                    })

cell_format2 = workbook.add_format({'font_color': 'black',
                                    'align': 'center',
                                    'valign': 'vcenter'
                                    })

cell_format3 = workbook.add_format({'bg_color': '#C0C0C0',
                                    'align': 'center',
                                    'valign': 'vcenter'
                                    })
p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')

p1.off(1)
time.sleep(1)
p1.on(5, 2, 1)
time.sleep(1)
a = SHSecureMode()
a.securemode(1)
# a.unlockDIMM(0)
# a.unlockIDT(0)
a.enable(1)
time.sleep(1)

worksheet.write(1, 0, 'Register', cell_format1)
worksheet.write(1, 1, 'Default', cell_format1)
worksheet.write(1, 2, '1st Write', cell_format1)
worksheet.write(1, 3, '1st Read', cell_format1)
worksheet.write(1, 4, '2nd Write', cell_format1)
worksheet.write(1, 5, '2nd Read', cell_format1)
worksheet.write(1, 6, '3rd Write', cell_format1)
worksheet.write(1, 7, '3rd Read', cell_format1)
worksheet.write(1, 8, '4th Write', cell_format1)
worksheet.write(1, 9, '4th Read', cell_format1)
worksheet.write(1, 10, 'Status', cell_format1)

# 0x00~0xFF
for i in range(256):
    data = data1 = data2 = data3 = data4 = 0
    ret, data = a.dsh.i2c_read_reg(i)
    if i == 0x34:
        j = 0
    if i != 0x34:
        a.dsh.i2c_write_reg(i, 0x00)
        ret, data1 = a.dsh.i2c_read_reg(i)
        a.dsh.i2c_write_reg(i, 0xFF)
        ret, data2 = a.dsh.i2c_read_reg(i)
        a.dsh.i2c_write_reg(i, 0x55)
        ret, data3 = a.dsh.i2c_read_reg(i)
        a.dsh.i2c_write_reg(i, 0xAA)
        ret, data4 = a.dsh.i2c_read_reg(i)
    if data1 is None:
        data1 = data2 = data3 = data4 = 'Null'
    else:
        data1 = hex(data1)
        data2 = hex(data2)
        data3 = hex(data3)
        data4 = hex(data4)
    data = hex(data)

    print(i, hex(i), data, data1, data2, data3, data4)

    if data1 == data2 == data3 == data4 == data:
        Status = 'Pass'
        cell_format = workbook.add_format({'font_color': 'green',
                                           'bold': True,
                                           'align': 'center',
                                           'valign': 'vcenter'
                                           })
    else:
        Status = 'Fail'
        cell_format = workbook.add_format({'font_color': 'red',
                                           'bold': True,
                                           'align': 'center',
                                           'valign': 'vcenter'
                                           })

    brow = 2
    worksheet.write(brow + i, 0, hex(i), cell_format2)
    worksheet.write(brow + i, 1, data, cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, data1, cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, data2, cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, data3, cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, data4, cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

workbook.close()
