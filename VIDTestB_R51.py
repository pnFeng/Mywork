import visa
import sys
import time
import xlsxwriter
import dolphin
from ICs import DolphinSH
from bit_manipulate import bm


class DMM:
    def __init__(self):
        rm = visa.ResourceManager()
        try:
            self.dmm1 = rm.open_resource('USB0::0x2A8D::0x1401::MY57215694::0::INSTR')
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

        # Diable OVP
        self.dsh.i2c_write_reg(0x56, 0xF0)

    def enable(self):
        # enable
        ret, data = self.dsh.i2c_read_reg(0x2F)
        print('data', data, 'ret', ret)
        data = data | 0x04
        self.dsh.i2c_write_reg(0x2F, data)
        self.dsh.i2c_write_reg(0x32, 0x80)
        self.dsh.i2c_write_reg(0x6B, data)

    def VIDSetting(self, ch, VID, cre):
        if ch == 'b':
            self.value = VID << 1 | cre
            self.dsh.i2c_write_reg(0x25, self.value)

    def ADCConfig(self, ch):
        if ch == 'a':
            self.dsh.i2c_write_reg(0x30, 0x80)
        elif ch == 'b':
            self.dsh.i2c_write_reg(0x30, 0x90)
        elif ch == 'c':
            self.dsh.i2c_write_reg(0x30, 0x98)

    def ADCrslt(self):
        ret, data1 = self.dsh.i2c_read_reg(0x31)
        ret, data2 = self.dsh.i2c_read_reg(0x72)
        return data1, data2

    def enableFLT(self, enable):
        if enable == 1:
            ret, data = self.dsh.i2c_read_reg(0xc7)
            data = data | 0x40
            self.dsh.i2c_write_reg(0xc7, data)
            ret, data = self.dsh.i2c_read_reg(0xc7)
            print('0xc7_1', hex(data))
        elif enable == 0:
            ret, data = self.dsh.i2c_read_reg(0xc7)
            data = data & 0xBF
            self.dsh.i2c_write_reg(0xc7, data)
            ret, data = self.dsh.i2c_read_reg(0xc7)
            print('0xc7_0', hex(data))

    def SetR51(self, bEn):
        if bEn == 1:
            ret, data = self.dsh.i2c_read_reg(0x51)
            assert ret == 0
            data = bm.set_bit(data, 3)
            self.dsh.i2c_write_reg(0x51, data)
        elif bEn == 0:
            ret, data = self.dsh.i2c_read_reg(0x51)
            assert ret == 0
            data = bm.clear_bit(data, 3)
            self.dsh.i2c_write_reg(0x51, data)
        ret, data = self.dsh.i2c_read_reg(0x51)
        print('0x51,{}'.format(hex(data)))

    def SetR6D(self, bEn):
        if bEn == 1:
            ret, data = self.dsh.i2c_read_reg(0x6D)
            assert ret == 0
            data = bm.set_bit(data, 1)
            self.dsh.i2c_write_reg(0x6D, data)
        elif bEn == 0:
            ret, data = self.dsh.i2c_read_reg(0x6D)
            assert ret == 0
            data = bm.clear_bit(data, 1)
            self.dsh.i2c_write_reg(0x6D, data)
        ret, data = self.dsh.i2c_read_reg(0x6D)
        print('0x6D,{}'.format(hex(data)))


workbook = xlsxwriter.Workbook('CHB_NL_R51_1.xlsx')
worksheet = workbook.add_worksheet()
chart1 = workbook.add_chart({'type': 'line'})
chart1.set_title({'name': ' Error(%) of Test Value Vs Target'})
chart1.set_x_axis({'name': 'Output Setting'})
chart1.set_y_axis({'name': 'Error(%)'})
chart1.set_size({'width': 1000, 'height': 600})

chart2 = workbook.add_chart({'type': 'line'})
chart2.set_title({'name': 'ADC Vs Multimeter Error(%)'})
chart2.set_x_axis({'name': 'Output Setting'})
chart2.set_y_axis({'name': 'Error(%)'})
chart2.set_size({'width': 1000, 'height': 600})

chart3 = workbook.add_chart({'type': 'line'})
chart3.set_title({'name': 'Test Value Vs Target'})
chart3.set_x_axis({'name': 'Output Setting'})
chart3.set_y_axis({'name': 'Test Value'})
chart3.set_size({'width': 1000, 'height': 600})

chart4 = workbook.add_chart({'type': 'line'})
chart4.set_title({'name': 'Delta_LSB'})
chart4.set_x_axis({'name': 'Output Setting'})
chart4.set_y_axis({'name': 'Delta_LSB'})
chart4.set_size({'width': 1000, 'height': 600})

m = DMM()
time.sleep(1)
a = SHVID()
a.ADCConfig('b')
a.enableFLT(1)
a.enable()
time.sleep(1)

for j in range(4):  # 10
    if j == 0:
        a.SetR51(0)
        a.SetR6D(0)
    elif j == 1:
        a.SetR6D(0)
        a.SetR51(1)
    elif j == 2:
        a.SetR6D(1)
        a.SetR51(0)
    elif j == 3:
        a.SetR6D(1)
        a.SetR51(1)

    localtime = time.asctime(time.localtime(time.time()))
    worksheet.write(1, 0 + 12 * j, '#')
    worksheet.write(1, 1 + 12 * j, 'Time')
    worksheet.write(1, 2 + 12 * j, 'Multimeter VOUT(V)')
    worksheet.write(1, 3 + 12 * j, 'Setting')
    worksheet.write(1, 4 + 12 * j, 'Readback')
    worksheet.write(1, 5 + 12 * j, 'ADC')
    worksheet.write(1, 6 + 12 * j, 'ADC Value(mV)')
    worksheet.write(1, 7 + 12 * j, 'Target(mV)')
    worksheet.write(1, 8 + 12 * j, 'Error(%)')
    worksheet.write(1, 9 + 12 * j, 'ADC Error(%)')
    worksheet.write(1, 10 + 12 * j, 'ADC Delta_LSB')

    n = 128
    for i in range(n):
        print(i, j)
        a.VIDSetting('b', i, 0)
        time.sleep(0.5)
        ret, data = a.dsh.i2c_read_reg(0x25)
        meas1 = m.meas()

        Sum = 0
        list1 = []
        for k in range(10):
            ADC1, ADC2 = a.ADCrslt()
            ADCData = (ADC1 * 15)
            Sum += ADC1
            time.sleep(0.05)
        ADCavg = Sum // 10

        targetADC_8bit = int(round(meas1 / 0.015))
        LSB_8bit = targetADC_8bit - ADC1

        if j == 0:
            target = 800 + 5 * i
        elif j == 1:
            target = 600 + 5 * i
        elif j == 2:
            target = 600 + 3.75 * i
        elif j == 3:
            target = 600 + 5 * i

        error = 100 * (1000 * meas1 - target) / target

        ADCData = (ADCavg * 15)
        ADCerror = 100 * (ADCData - 1000 * meas1) / (1000 * meas1)

        worksheet.write(i + 2, 0 + 12 * j, i)
        worksheet.write(i + 2, 1 + 12 * j, localtime)
        worksheet.write(i + 2, 2 + 12 * j, meas1)
        worksheet.write(i + 2, 3 + 12 * j, a.value)
        worksheet.write(i + 2, 4 + 12 * j, hex(data))
        worksheet.write(i + 2, 5 + 12 * j, hex(ADCavg))
        worksheet.write(i + 2, 6 + 12 * j, ADCData)
        worksheet.write(i + 2, 7 + 12 * j, target)
        worksheet.write(i + 2, 8 + 12 * j, error)
        worksheet.write(i + 2, 9 + 12 * j, ADCerror)
        worksheet.write(i + 2, 10 + 12 * j, LSB_8bit)

    s = 'range_' + str(j)
    cell1 = xlsxwriter.utility.xl_rowcol_to_cell(2, 8 + 12 * j)
    cell2 = xlsxwriter.utility.xl_rowcol_to_cell(129, 8 + 12 * j)
    s1 = '=Sheet1!' + cell1 + ':' + cell2
    chart1.add_series({
        'name': s,
        'categories': '=Sheet1!E2:E130',
        'values': s1,
    })

    cell1 = xlsxwriter.utility.xl_rowcol_to_cell(2, 9 + 12 * j)
    cell2 = xlsxwriter.utility.xl_rowcol_to_cell(129, 9 + 12 * j)
    s2 = '=Sheet1!' + cell1 + ':' + cell2
    chart2.add_series({
        'name': s,
        'categories': '=Sheet1!E2:E130',
        'values': s2,
    })

    cell1 = xlsxwriter.utility.xl_rowcol_to_cell(2, 2 + 12 * j)
    cell2 = xlsxwriter.utility.xl_rowcol_to_cell(129, 2 + 12 * j)
    s3 = '=Sheet1!' + cell1 + ':' + cell2
    chart3.add_series({
        'name': s,
        'categories': '=Sheet1!E2:E130',
        'values': s3,
    })

    cell1 = xlsxwriter.utility.xl_rowcol_to_cell(2, 10 + 12 * j)
    cell2 = xlsxwriter.utility.xl_rowcol_to_cell(129, 10 + 12 * j)
    s4 = '=Sheet1!' + cell1 + ':' + cell2
    chart4.add_series({
        'name': s,
        'categories': '=Sheet1!E2:E130',
        'values': s4,
        'line': {
            'width': 0.25,
        },
    })

worksheet.insert_chart('A135', chart1)
worksheet.insert_chart('R135', chart2)
worksheet.insert_chart('AI135', chart3)
worksheet.insert_chart('AY135', chart4)

workbook.close()
