# -*- coding: utf-8 -*-

import sys
import time
import xlsxwriter
# import dolphin
from ICs import DolphinSH


class SH_Unlock:
    def __init__(self):
        self.dsh = DolphinSH()
        ret = self.dsh.connect()
        if ret is False:
            print("Failed to connect")
            sys.exit(-1)

        # self.dsh.helper_mark_global_vars(lid=self.dsh.LID, hid=7)
        self.dsh.set_voltage(1.8)
        self.dsh.set_i2c_bitrate(100)

        # Erase R04 - R07
        #self.dsh.i2c_write_reg(0x39, 0x74)
        #time.sleep(0.5)

        # Un-lock DIMM vendor region(0x40~0x7f)
        self.dsh.i2c_write_reg(0x37, 0x73)
        self.dsh.i2c_write_reg(0x38, 0x94)
        self.dsh.i2c_write_reg(0x39, 0x40)
        #
        # # Un-lock PMIC vendor region(0x80~0xff)
        self.dsh.i2c_write_reg(0x37, 0x79)
        self.dsh.i2c_write_reg(0x38, 0xbe)
        self.dsh.i2c_write_reg(0x39, 0x10)
        #
        # # Diable OVP
        # self.dsh.i2c_write_reg(0x56, 0xF0)

    def VR_Dis(self):
        # disable
        ret, data = self.dsh.i2c_read_reg(0x2F)
        print('data', data, 'ret', ret)
        data = data & 0xFB
        self.dsh.i2c_write_reg(0x2F, data)
        self.dsh.i2c_write_reg(0x32, 0x80)


workbook = xlsxwriter.Workbook('../Host_DIMM_Region_RV_Registers_Test.xlsx')
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

a = SH_Unlock()
a.VR_Dis()
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

# 0x00~0x1F
for i in range(16):
    ret, data = a.dsh.i2c_read_reg(i)
    a.dsh.i2c_write_reg(i, 0x00)
    ret, data1 = a.dsh.i2c_read_reg(i)
    a.dsh.i2c_write_reg(i, 0xFF)
    ret, data2 = a.dsh.i2c_read_reg(i)
    a.dsh.i2c_write_reg(i, 0x55)
    ret, data3 = a.dsh.i2c_read_reg(i)
    a.dsh.i2c_write_reg(i, 0xAA)
    ret, data4 = a.dsh.i2c_read_reg(i)
    if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
        workbook.close()
        raise ValueError("data error")
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

    print(hex(i), data, data1, data2, data3, data4)

    brow = 2
    worksheet.write(brow + i, 0, hex(i), cell_format2)
    worksheet.write(brow + i, 1, hex(data), cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, hex(data1), cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, hex(data2), cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, hex(data3), cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, hex(data4), cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

# 0x10~0x14
for i in range(5):
    ret, data = a.dsh.i2c_read_reg(i + 16)
    a.dsh.i2c_write_reg(i + 16, 0x00)
    ret, data1 = a.dsh.i2c_read_reg(i + 16)
    a.dsh.i2c_write_reg(i + 16, 0xFF)
    ret, data2 = a.dsh.i2c_read_reg(i + 16)
    a.dsh.i2c_write_reg(i + 16, 0x55)
    ret, data3 = a.dsh.i2c_read_reg(i + 16)
    a.dsh.i2c_write_reg(i + 16, 0xAA)
    ret, data4 = a.dsh.i2c_read_reg(i + 16)
    if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
        workbook.close()
        raise ValueError("data error")
    if data1 == data2 == data3 == data4 == data == 0x00:
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

    print(hex(i + 16), data, data1, data2, data3, data4)

    brow = 19
    worksheet.write(brow + i, 0, hex(i + 16), cell_format2)
    worksheet.write(brow + i, 1, hex(data), cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, hex(data1), cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, hex(data2), cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, hex(data3), cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, hex(data4), cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

# 0x15~0x31
for i in range(29):
    ret, data = a.dsh.i2c_read_reg(i + 21)
    a.dsh.i2c_write_reg(i + 21, 0x00)
    ret, data1 = a.dsh.i2c_read_reg(i + 21)
    a.dsh.i2c_write_reg(i + 21, 0xFF)
    ret, data2 = a.dsh.i2c_read_reg(i + 21)
    a.dsh.i2c_write_reg(i + 21, 0x55)
    ret, data3 = a.dsh.i2c_read_reg(i + 21)
    a.dsh.i2c_write_reg(i + 21, 0xAA)
    ret, data4 = a.dsh.i2c_read_reg(i + 21)
    if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
        workbook.close()
        raise ValueError("data error")
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

    print(hex(i + 21), data, data1, data2, data3, data4)

    brow = 25
    worksheet.write(brow + i, 0, hex(i + 21), cell_format2)
    worksheet.write(brow + i, 1, hex(data), cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, hex(data1), cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, hex(data2), cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, hex(data3), cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, hex(data4), cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

# 0x32
ret, data = a.dsh.i2c_read_reg(0x32)
a.dsh.i2c_write_reg(0x32, 0x00)
ret, data1 = a.dsh.i2c_read_reg(0x32)
a.dsh.i2c_write_reg(0x32, 0x3F)
ret, data2 = a.dsh.i2c_read_reg(0x32)
a.dsh.i2c_write_reg(0x32, 0x15)
ret, data3 = a.dsh.i2c_read_reg(0x32)
a.dsh.i2c_write_reg(0x32, 0x2A)
ret, data4 = a.dsh.i2c_read_reg(0x32)
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

print('0x32', data, data1, data2, data3, data4)

brow = 55
worksheet.write(brow, 0, hex(0x32), cell_format2)
worksheet.write(brow, 1, hex(data), cell_format2)
worksheet.write(brow, 2, '0x00', cell_format3)
worksheet.write(brow, 3, hex(data1), cell_format2)
worksheet.write(brow, 4, '0x3f', cell_format3)
worksheet.write(brow, 5, hex(data2), cell_format2)
worksheet.write(brow, 6, '0x15', cell_format3)
worksheet.write(brow, 7, hex(data3), cell_format2)
worksheet.write(brow, 8, '0x2a', cell_format3)
worksheet.write(brow, 9, hex(data4), cell_format2)
worksheet.write(brow, 10, Status, cell_format)

# 0x33
ret, data = a.dsh.i2c_read_reg(0x33)
a.dsh.i2c_write_reg(0x33, 0x00)
ret, data1 = a.dsh.i2c_read_reg(0x33)
a.dsh.i2c_write_reg(0x33, 0xFF)
ret, data2 = a.dsh.i2c_read_reg(0x33)
a.dsh.i2c_write_reg(0x33, 0x55)
ret, data3 = a.dsh.i2c_read_reg(0x33)
a.dsh.i2c_write_reg(0x33, 0xAA)
ret, data4 = a.dsh.i2c_read_reg(0x33)
if data1 == data2 == data3 == data4 == data == 0x00:
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

print('0x33', data, data1, data2, data3, data4)

brow = 57
worksheet.write(brow, 0, hex(0x33), cell_format2)
worksheet.write(brow, 1, hex(data), cell_format2)
worksheet.write(brow, 2, '0x00', cell_format3)
worksheet.write(brow, 3, hex(data1), cell_format2)
worksheet.write(brow, 4, '0xff', cell_format3)
worksheet.write(brow, 5, hex(data2), cell_format2)
worksheet.write(brow, 6, '0x55', cell_format3)
worksheet.write(brow, 7, hex(data3), cell_format2)
worksheet.write(brow, 8, '0xaa', cell_format3)
worksheet.write(brow, 9, hex(data4), cell_format2)
worksheet.write(brow, 10, Status, cell_format)

# 0x34
ret, data = a.dsh.i2c_read_reg(0x34)
a.dsh.i2c_write_reg(0x34, 0x0E)
ret, data1 = a.dsh.i2c_read_reg(0x34)
a.dsh.i2c_write_reg(0x34, 0x0F)
ret, data2 = a.dsh.i2c_read_reg(0x34)
a.dsh.i2c_write_reg(0x34, 0x1E)
ret, data3 = a.dsh.i2c_read_reg(0x34)
a.dsh.i2c_write_reg(0x34, 0x1F)
ret, data4 = a.dsh.i2c_read_reg(0x34)
if data1 == data2 == data3 == data4 == data == 0x0E:
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

print('0x34', data, data1, data2, data3, data4)

brow = 59
worksheet.write(brow, 0, hex(0x34), cell_format2)
worksheet.write(brow, 1, hex(data), cell_format2)
worksheet.write(brow, 2, '0x0E', cell_format3)
worksheet.write(brow, 3, hex(data1), cell_format2)
worksheet.write(brow, 4, '0x0f', cell_format3)
worksheet.write(brow, 5, hex(data2), cell_format2)
worksheet.write(brow, 6, '0x1E', cell_format3)
worksheet.write(brow, 7, hex(data3), cell_format2)
worksheet.write(brow, 8, '0x1F', cell_format3)
worksheet.write(brow, 9, hex(data4), cell_format2)
worksheet.write(brow, 10, Status, cell_format)

# 0x35~0x3f
for i in range(11):
    ret, data = a.dsh.i2c_read_reg(i + 53)
    a.dsh.i2c_write_reg(i + 53, 0x00)
    ret, data1 = a.dsh.i2c_read_reg(i + 53)
    a.dsh.i2c_write_reg(i + 53, 0xFF)
    ret, data2 = a.dsh.i2c_read_reg(i + 53)
    a.dsh.i2c_write_reg(i + 53, 0x55)
    ret, data3 = a.dsh.i2c_read_reg(i + 53)
    a.dsh.i2c_write_reg(i + 53, 0xAA)
    ret, data4 = a.dsh.i2c_read_reg(i + 53)
    if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
        workbook.close()
        raise ValueError("data error")
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

    print(hex(i + 53), data, data1, data2, data3, data4)

    brow = 61
    worksheet.write(brow + i, 0, hex(i + 53), cell_format2)
    worksheet.write(brow + i, 1, hex(data), cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, hex(data1), cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, hex(data2), cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, hex(data3), cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, hex(data4), cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

# 0x40~0x6F
for i in range(48):
    ret, data = a.dsh.i2c_read_reg(i + 64)
    a.dsh.i2c_write_reg(i + 64, 0x00)
    ret, data1 = a.dsh.i2c_read_reg(i + 64)
    a.dsh.i2c_write_reg(i + 64, 0xFF)
    ret, data2 = a.dsh.i2c_read_reg(i + 64)
    a.dsh.i2c_write_reg(i + 64, 0x55)
    ret, data3 = a.dsh.i2c_read_reg(i + 64)
    a.dsh.i2c_write_reg(i + 64, 0xAA)
    ret, data4 = a.dsh.i2c_read_reg(i + 64)
    if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
        workbook.close()
        raise ValueError("data error")
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

    print(hex(i + 64), data, data1, data2, data3, data4)

    brow = 73
    worksheet.write(brow + i, 0, hex(i + 64), cell_format2)
    worksheet.write(brow + i, 1, hex(data), cell_format2)
    worksheet.write(brow + i, 2, '0x00', cell_format3)
    worksheet.write(brow + i, 3, hex(data1), cell_format2)
    worksheet.write(brow + i, 4, '0xff', cell_format3)
    worksheet.write(brow + i, 5, hex(data2), cell_format2)
    worksheet.write(brow + i, 6, '0x55', cell_format3)
    worksheet.write(brow + i, 7, hex(data3), cell_format2)
    worksheet.write(brow + i, 8, '0xaa', cell_format3)
    worksheet.write(brow + i, 9, hex(data4), cell_format2)
    worksheet.write(brow + i, 10, Status, cell_format)

# 0x70~0xFF
for i in range(144):
    if i != 18 or i != 83:
        ret, data = a.dsh.i2c_read_reg(i + 112)
        a.dsh.i2c_write_reg(i + 112, 0x00)
        ret, data1 = a.dsh.i2c_read_reg(i + 112)
        a.dsh.i2c_write_reg(i + 112, 0xFF)
        ret, data2 = a.dsh.i2c_read_reg(i + 112)
        a.dsh.i2c_write_reg(i + 112, 0x55)
        ret, data3 = a.dsh.i2c_read_reg(i + 112)
        a.dsh.i2c_write_reg(i + 112, 0xAA)
        ret, data4 = a.dsh.i2c_read_reg(i + 112)
        if not (isinstance(data, int) and isinstance(data1, int) and isinstance(data2, int) and isinstance(data3, int)):
            workbook.close()
            raise ValueError("data error")
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

        print(hex(i + 112), data, data1, data2, data3, data4)

        brow = 122
        worksheet.write(brow + i, 0, hex(i + 112), cell_format2)
        worksheet.write(brow + i, 1, hex(data), cell_format2)
        worksheet.write(brow + i, 2, '0x00', cell_format3)
        worksheet.write(brow + i, 3, hex(data1), cell_format2)
        worksheet.write(brow + i, 4, '0xff', cell_format3)
        worksheet.write(brow + i, 5, hex(data2), cell_format2)
        worksheet.write(brow + i, 6, '0x55', cell_format3)
        worksheet.write(brow + i, 7, hex(data3), cell_format2)
        worksheet.write(brow + i, 8, '0xaa', cell_format3)
        worksheet.write(brow + i, 9, hex(data4), cell_format2)
        worksheet.write(brow + i, 10, Status, cell_format)

workbook.close()
