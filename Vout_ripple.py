# -*- coding: UTF-8 -*-
import time
from xlsxwriter.workbook import Workbook

import idt_global
from idt_errors import TargetException
import idt_errors
import sys
import os
import traceback

import ics.PMIC.P8911 as PH

CASE_TEMPLATE_VER = 0.1

HELP_FILE = ""
PARAM_FILE = "scope_test.cfg"

PARAM_LIST = [{'name': 'Oscilloscope', 'type': 'Oscilloscope', 'model': 'HDO8180'},
              {'name': 'Vin_Bulk', 'type': 'Power', 'model': 'E3631A', 'para': [['Volt(V)', []],
                                                                              ['Current Limit(A)', 3]]},

              {'name': 'E-Load', 'type': 'eload', 'model': '6314A', 'para': [['Values', []]]},
              {'name': 'I2C Adapter', 'type': 'I2C Adapter', 'model': 'Aardvark'},

              {'name': 'Configuration', 'type': 'Customized', 'para': [['Frequcency', '750K|1000K|1250K|1500K'],
                                                                       ['Mode', 'CCM|DCM']]},
              {'name': 'Select Load', 'type': 'Customized', 'para': [['Load Table', 'Corvette|Viper']]},

              {'name': 'File Name', 'type': 'Customized', 'para': [['File name keyword', '']]},

              {'name': 'Channel Enable/Disable', 'type': 'Customized', 'para': [['Channel AB', 'unchecked/checked'],
                                                                                ['Channel A', 'unchecked/checked'],
                                                                                ['Channel B', 'unchecked/checked'],
                                                                                ['Channel C', 'unchecked/checked'],
                                                                                ['Channel All OFF','unchecked/checked']]},

              {'name': 'Scope configure', 'type': 'Customized', 'para': [['Noise Filter(+__bits)', '0|0.5|1|1.5|2|2.5|3'],
                                                                         ['Inductor(uH)', []],
                                                                         ['Minimum Trigger Number', '50|100|200|500|1000'],
                                                                         ['Maximum Wait Time(s)', '5|10|20|40|80|120'],
                                                                         ['MaxTimebase(us)', ''],
                                                                         ['MinTimebase(us)', '']]}]


def __time_stamp(logging, prefix_str):
    h = time.localtime().tm_hour
    m = time.localtime().tm_min
    s = time.localtime().tm_sec
    logging.info('{} Time: {:4d}{:02d}{:02d}-{:02d}{:02d}{:02d}'.format(
        prefix_str,
        time.localtime().tm_year,
        time.localtime().tm_mon,
        time.localtime().tm_mday,
        h, m, s))
    t = h * 3600 + m * 60 + s
    #print t
    return t

def auto_horizontal(MaxTimebase, MinTimebase, Mode, Iload, Ipp, Fsw, Nphase):
    Timebase = MaxTimebase
    if Iload == 0:
        Timebase = MaxTimebase
        return Timebase
    if Mode == 'DCM':
        print('++++++++++++++DCM +++++++++++++++++++++')
        if Iload < Ipp * Nphase:
            Timebase = 2 * (Nphase * Ipp / (Iload * Fsw * 2.0))
        else:
            Timebase = 4 / Fsw  # CCM 4 period for each division
    if Mode == 'CCM':
        print('++++++++++++++CCM +++++++++++++++++++++')
        T = 1 / float(Fsw)
        Timebase = 4 * T
        print(Timebase)
    # clamp Timebase between MaxTimebase and MinTimebase
    if Timebase > MaxTimebase:
        Timebase = MaxTimebase
    elif Timebase < MinTimebase:
        Timebase = MinTimebase
    return Timebase


def auto_horizontal_1(MaxTimebase, MinTimebase, Mode, Iload, Ipp, Fsw, Nphase):
    T = 1 / float(Fsw)
    if (Mode == 'CCM') or (Iload >= Ipp * 0.5):  # PMIC must in the CCM
        Timebase = 4 * T                         # each division will have 2 period waveform
    elif Iload == 0:                             # PMIC must in the DCM
        Timebase = MaxTimebase
        return Timebase
    else:
        Ti  = Nphase * Ipp / (Iload * Fsw * 2.0) # Ti( ripple period in DCM) is the function of Iload, Ti = f(Iload).
        Timebase = 4 * Ti # each division will have 2 period waveform

    if Timebase > MaxTimebase: # clamp Timebase between MaxTimebase and MinTimebase
        Timebase = MaxTimebase
    elif Timebase < MinTimebase:
        Timebase = MinTimebase
    return Timebase


def execute(env, logging):
    print('pro path:{}'.format(idt_global.logsdir))
    t1 = __time_stamp(logging, 'Start')

    powerVinBulk = env.get('Vin_Bulk')
    list_VinBulk = powerVinBulk.para['Volt(V)']
    VinBulk_CL = powerVinBulk.para['Current Limit(A)']

    eload = env.get('E-Load')
    list_load = eload.para['Values']  # floatrange(el.para['Min'], el.para['Max'], el.para['step'])
    if not isinstance(list_load, list):
        logging.log('Parameters of E-Load is not a list!')
        return idt_errors.PATN_PARAM_ERROR

    i2c = env.get('I2C Adapter')
    PH.set_i2c_device(i2c)
    file_name_keyword = env.get('File Name').para['File name keyword']
    config = env.get('Configuration')
    Fre = config.para['Frequcency']
    Mode = config.para['Mode']
    load_table = env.get('Select Load').para['Load Table']

    Scope_configure = env.get('Scope configure')
    filter_bit = Scope_configure.para['Noise Filter(+__bits)']
    Inductance = Scope_configure.para['Inductor(uH)']
    Minimum_Trigger_Number = Scope_configure.para['Minimum Trigger Number']
    Minimum_Trigger_Number = int(Minimum_Trigger_Number)

    Maximum_Wait_Time = Scope_configure.para['Maximum Wait Time(s)']
    Maximum_Wait_Time = int(Maximum_Wait_Time)

    MaxTimebase = Scope_configure.para['MaxTimebase(us)']
    MinTimebase = Scope_configure.para['MinTimebase(us)']
    # MaxTimebase = 10e-3  # max time base at 0A set as 10ms
    MaxTimebase = MaxTimebase * 1e-6
    MinTimebase = MinTimebase * 1e-6

    ch_ed = env.get('Channel Enable/Disable')

    osc = env.get('Oscilloscope')
    # step 1 set up the scope. Channel
    osc.set_save_scale('50')
    osc.set_channel_filter(filter_bit)
    osc.set_channel_label('Vout')
    # step 2 set up the measurement in scope
    osc.set_measurement('p1', 'PeakToPeak')
    osc.set_measurement('p2', 'Maximum')
    osc.set_measurement('p3', 'Minimum')
    osc.set_measurement('p4', 'Mean')
    osc.set_measurement('p5', 'Frequency')
    osc.set_measurement('p6', 'Period')


    ts_row = 0
    sc_row = 0
    print('')
    selected_rail = ''
    Vout_Normal = 1.1
    Nphase = 1
    if ch_ed.para['Channel AB'] == 'checked':
        selected_rail = 'AB'
        Vout_Normal = 1.1
        Nphase = 2
    elif ch_ed.para['Channel A'] == 'checked':
        selected_rail = 'A'
        Vout_Normal = 1.1
        Nphase = 1
    elif ch_ed.para['Channel B'] == 'checked':
        selected_rail = 'B'
        Vout_Normal = 1.1
        Nphase = 1
    elif ch_ed.para['Channel C'] == 'checked':
        selected_rail = 'C'
        Vout_Normal = 1.8
        Nphase = 1
    elif ch_ed.para['Channel D'] == 'checked':
        selected_rail = 'D'
        Vout_Normal = 1.8
        Nphase = 1
    elif ch_ed.para['Channel All OFF'] == 'checked':
        selected_rail = 'all_off'

    print('Rail: {} , Nphase: {} ,Vout: {}V , Frequcency: {} ,Mode: {} ,Inductor: {}uH'.\
        format(selected_rail, Nphase,Vout_Normal, Fre, Mode, Inductance ))

    # Open the record xlsx file
    pro_path = idt_global.logsdir
    resultpath = os.path.join(pro_path, "result")
    if not os.path.exists(resultpath):
        os.makedirs(resultpath)
    exlbook = Workbook('{}\\RIPPLE_{}_{}_{}uH_{}_{}_{:4d}{:02d}{:02d}-{:02d}{:02d}{:02d}.xlsx'.format(
        resultpath, selected_rail, file_name_keyword, Inductance, Fre, Mode,
        time.localtime().tm_year, time.localtime().tm_mon, time.localtime().tm_mday,
        time.localtime().tm_hour, time.localtime().tm_min, time.localtime().tm_sec))
    datasheet_name = 'Data'
    scsheet_name = 'Waveforms'
    datasheet = exlbook.add_worksheet(datasheet_name)
    scsheet = exlbook.add_worksheet(scsheet_name)

    if not list_load:
        print('.........Load List unspecified, loading default value list...........')
        if load_table == 'Corvette':
            list_load = PH.LOAD_TABLE_Ripple_Corvette[selected_rail]
        elif load_table == 'Viper':
            list_load = PH.LOAD_TABLE_Ripple_Viper[selected_rail]
        print(list_load)
        if type(list_load) == type('a'):
            if list_load.find('min') == 0:
                list_load = PH.param_change(list_load)
            elif list_load.find('min') == -1:
                list_load = list_load.split(',')
        # list_load.reverse()
        print(list_load)


    eload.setLoadState('OFF')
    time.sleep(1)
    powerVinBulk.off()
    time.sleep(1)

    powerVinBulk.set_current_limit(VinBulk_CL)
    powerVinBulk.set_voltage(PH.Defalt_Vin_Bulk)
    powerVinBulk.on()
    time.sleep(1)
    PH.write(0x2F, 0x06)  # Disable All rail and set to programmable mode




    ######## set Frequency ###################################################
    reg29 = PH.read(0x29)
    reg2A = PH.read(0x2A)
    if Fre == '750K':
        print('--------Frequency is 750K---------')
        reg29 = reg29 & 0b11001100
        reg2A = reg2A & 0b11001100
    elif Fre == '1000K':
        print('--------Frequency is 1000K---------')
        reg29 = reg29 & 0b11011101 | 0b00010001
        reg2A = reg2A & 0b11011101 | 0b00010001
    elif Fre == '1250K':
        print('--------Frequency is 1250K---------')
        reg29 = reg29 & 0b11101110 | 0b00100010
        reg2A = reg2A & 0b11101110 | 0b00100010
    elif Fre == '1500K':
        print('--------Frequency is 1500K---------')
        reg29 = reg29 | 0b00110011
        reg2A = reg2A | 0b00110011
    PH.write(0x29, reg29)
    PH.write(0x2A, reg2A)
    #########################################################################

    ######## set Mode #######################################################
    reg29 = PH.read(0x29)
    reg2A = PH.read(0x2A)
    if Mode == 'DCM':
        print('--------Mode is DCM---------')
        reg29 = reg29 & 0b10111011 | 0b10001000
        reg2A = reg2A & 0b10111011 | 0b10001000
    elif Mode == 'CCM':
        print('--------Mode is CCM---------')
        reg29 = reg29 | 0b11001100
        reg2A = reg2A | 0b11001100
    PH.write(0x29, reg29)
    PH.write(0x2A, reg2A)
    reg29 = PH.read(0x29)
    reg2A = PH.read(0x2A)
    print('0x29 = {}, 0x2A = {}'.format(bin(reg29), bin(reg2A)))
    ###########################################################################

    ################## enable channel #########################################
    # if ch_ed.para['Channel AB'] == 'checked':
    #     PH.enable_buck(PH.BUCK_AB)
    # elif ch_ed.para['Channel A'] == 'checked':
    #     PH.enable_buck(PH.BUCK_A)
    # elif ch_ed.para['Channel B'] == 'checked':
    #     PH.enable_buck(PH.BUCK_B)
    # elif ch_ed.para['Channel C'] == 'checked':
    #     PH.enable_buck(PH.BUCK_C)
    # else:
    #     logging.log('[ERROR] PMEC channel all disable!')
    #     return idt_errors.PATN_PARAM_ERROR
    # time.sleep(2)
    ############################################################################
    ######## enable channel #################################
    if ch_ed.para['Channel AB'] == 'checked':
        print('--------Dual_Phase(A B)---------')
        PH.Dual_Phase()
        PH.enable_buck(PH.BUCK_AB)
    elif ch_ed.para['Channel A'] == 'checked':
        print('--------Single_Phase(A)---------')
        PH.enable_buck_a()
    elif ch_ed.para['Channel B'] == 'checked':
        print('--------Single_Phase(B)---------')
        PH.enable_buck_b()
    elif ch_ed.para['Channel C'] == 'checked':
        print('--------Enable Rail (C)---------')
        PH.unlock_idt_region()
        aa = PH.read(0xB2)
        aa = aa & 0b11111110
        PH.write(0xB2, aa)

        PH.enable_buck_c()
    elif ch_ed.para['Channel All OFF'] == 'checked':
        print('--------All Disable---------')
        PH.enable_buck(PH.BUCK_NONE)
    else:
        logging.log('[ERROR] PMEC channel all disable!')
        return idt_errors.PATN_PARAM_ERROR
    #######################################################


    property1 = {
        'font_size': 15,  #
        # 'bold': True,  #
        'align': 'left',  #
        'valign': 'vcenter',  #
        'font_name': 'Microsoft YaHei UI',
        'border': 0,  #
        'text_wrap': False,  #
    }
    property2 = {
        'align': 'left',  #
        'valign': 'vcenter',  #
        'border': 2,  #
    }
    property3 = {
        'font_size': 13,  #
        # 'bold': True,  #
        'align': 'center',  #
        'valign': 'vcenter',  #
        'font_name': 'Microsoft YaHei UI',
        'border': 0,  #
        'text_wrap': False,  #
    }
    cell_format = exlbook.add_format(property1)
    merge_format = exlbook.add_format(property2)
    data_format = exlbook.add_format(property3)

    img_format = {
        'x_offset': 20,  #
        'y_offset': 10,  #
        'x_scale': 0.9,  #
        'y_scale': 1.1,  #
        'url': None,
        'tip': None,
        'image_data': None,
        'positioning': None
    }
    series_name_position_list = []
    datasheet.set_column(0, 10, 20)



    for Vin in list_VinBulk:
        print('Vin is {}V'.format(Vin))
        powerVinBulk.set_voltage(Vin)
        time.sleep(2)
        if eload.getVolt() < 0.5 * Vout_Normal :
            eload.setLoadState('OFF')
            time.sleep(1)
            PH.write(0x14, 0x01)  # Global clear

            if ch_ed.para['Channel A'] == 'checked':
                print ('--------Restart Rail (A)---------')
                PH.enable_buck_a()
            elif ch_ed.para['Channel B'] == 'checked':
                print ('--------Restart Rail (B)---------')
                PH.enable_buck_b()
            elif ch_ed.para['Channel C'] == 'checked':
                print ('--------Restart Rail (C)---------')
                PH.enable_buck_c()
        time.sleep(1)

        testnum = 1
        table_voltage = ('Vin = ' + str(Vin) + 'V', '')
        datasheet.write_row(ts_row, 1, table_voltage, data_format)
        series_name_position_list.append(ts_row)
        ts_row += 1
        table_head = ('Iout(A)', 'Vmin(V)', 'Vmax(V)', 'Vavg(V)', 'Vpp_max(mV)', 'Test number')
        datasheet.write_row(ts_row, 1, table_head, data_format)
        ts_row += 1


        for I in list_load:
            print ('                                                                                                    ')
            print ('----------------------------------------------------------------------------------------------------')
            I = float(I)
            if I <= 0.2:
                eload.setCCL(I)
                osc.set_coupling('DC1M')
                print ('Load : {}A , Coupling : DC1M'.format(I))
            elif I > 0.2 and I <= 2:
                eload.setCCM(I)
                osc.set_coupling('DC50')
                print ('Load : {}A , Coupling : DC50'.format(I))

            else:
                eload.setCCH(I)
                osc.set_coupling('DC50')
            eload.setLoadState('ON')
            time.sleep(2)

            # MaxTimebase = 10e-3  # max time base at 0A set as 10ms
            # MaxTimebase = 50e-6  # max time base at 0A set as 50us
            # MinTimebase = 5e-6  # min time base set as 5us

            if 'K' in Fre:
                Fre = Fre[0:-1]
            Fsw = int(Fre) * 10 ** 3
            L = float(Inductance) * 10 ** (-6)
            Ipp = (Vin - Vout_Normal) * Vout_Normal / (Vin * Fsw * L)
            time_base = auto_horizontal_1(MaxTimebase, MinTimebase, Mode, I, Ipp, Fsw, Nphase)
            print ('Timebase : {:.2f}uS / div'.format(time_base * 1e6))


            print ('---------------------- 1.set_trigger_mode is auto --------------------------')
            osc.set_trigger_mode('auto')
            print ('---------------------- 2.set_horizontal_offset is 0us --------------------------')
            osc.set_horizontal_offset(0)
            print ('---------------------- 3.Setting horizontal scale --------------------------')
            osc.set_horizontal(time_base)
            print ('---------------------- 4.Finding vertical scale --------------------------')

            osc.find_scale()
            if I <= 0.2:
                time.sleep(7)
            else:
                time.sleep(1)
            print ('---------------------- 5.trigger_find_level  --------------------------')
            osc.trigger_find_level()
            time.sleep(1)
            Vout_max = osc.get_value('p2', 'min')
            print ('---------------------- 6.set_trigger_level is Vout_max  --------------------------')
            osc.set_trigger_level(Vout_max)
            # print ('---------------------- 7.set_trigger_mode is normal  --------------------------')
            # osc.set_trigger_mode('normal')
            print ('---------------------- 8.set_label  --------------------------')
            label = 'Ripple ' + selected_rail + ' :  ' + str(Vin) + 'V  ' + str(I) + 'A'
            osc.set_channel_label(label)
            print ('---------------------- 9.set_label_position  --------------------------')
            osc.set_channel_label_position(0)

            time.sleep(1)
            osc.clear_Sweeps()


            trigger_number = 0
            wait_time = 0  # read trigger times every seconds
            while trigger_number <= Minimum_Trigger_Number and wait_time <= Maximum_Wait_Time:
                trigger_number = int(osc.get_value('p1', 'num'))
                # print 'trigger number is {},  {}s passed'.format(trigger_number,wait_time)
                wait_time = wait_time + 1
                time.sleep(1)
            print ('trigger number is {},  {}s passed'.format(trigger_number, wait_time))

            # ----------------Getting Data Start---------------------
            print ('---------------- Start Getting Data ---------------------')
            Vpp_max = round(float(osc.get_value('p1', 'max')) * 1000, 2)
            Vout_max = round(float(osc.get_value('p2', 'max')), 5)
            Vout_min = round(float(osc.get_value('p3', 'min')), 5)
            Vout_mean = round(float(osc.get_value('p4', 'mean')), 5)
            Period = float(osc.get_value('p6', 'mean'))
            print ('Load: {}A ,Min:{}V , Max:{}V ,Mean: {}, Vpp: {}mV,'.format(I, Vout_min, Vout_max, Vout_mean, Vpp_max))
            table_content = (I, Vout_min, Vout_max, Vout_mean, Vpp_max, testnum)
            datasheet.write_row(ts_row, 1, table_content, data_format)
            ts_row += 1
            print ('---------------- Getting Data Ended ---------------------')
            # ----------------Getting Data End---------------------

            # ----------capture screen----------------
            # screen_capture(osc, I, Vin, scsheet, testnum, sc_row)
            ripplepath = os.path.join(pro_path, "ripple")
            if not os.path.exists(ripplepath):
                os.makedirs(ripplepath)
            print ('---------------- Getting Waveform ---------------------')
            osc.StoreHardcopyToFile('PNG', '', ripplepath + '\Ripple' + str(Vin) +'V' + str(I)+'A' + '.PNG')

            data = ('Test#', testnum, 'Vin = ', str(Vin) + 'V', 'Iload = ', str(I) + 'A')
            scsheet.write_row(sc_row, 1, data, cell_format)
            scsheet.set_row(sc_row, 40)

            sc_row += 1
            scsheet.merge_range(sc_row, 1, sc_row + 44, 19, None, merge_format)
            print ('---------------- Inserting Waveform ---------------------')
            scsheet.insert_image('B' + str(sc_row + 1), ripplepath + '\Ripple' + str(Vin) +'V' + str(I)+'A' + '.PNG', img_format)
            sc_row += 46
            testnum += 1
            print ('----------------------------------------------------------------------------------------------------')
            print (' ')

        eload.setLoadState('OFF')
        time.sleep(1)
        ts_row += 1
    powerVinBulk.off()
    # osc.close()

    # Color name Vs RGB color code
    blue1 = '#4F81BD'
    red1 = '#C0504D'
    olive_green1 = '#9BBB59'
    color_list = [blue1, red1, olive_green1]

    chart = exlbook.add_chart({'type': 'scatter', 'subtype': 'smooth_with_markers'})
    chart.set_size({'width': 1000, 'height': 700})
    chart.set_title({'name': 'Ripple ( Rail ' + selected_rail + ' )', 'name_font': {'size': 20, 'bold': True, 'name': 'Arial'}})
    chart.set_x_axis({'name': 'Load Current(A)',
                      'name_font': {'size': 16, 'bold': True, 'name': 'Arial'},
                      'num_font': {'size': 12, 'italic': False, 'name': 'Microsoft YaHei UI'}})
    chart.set_y_axis({'name': 'Vpp_max(mV)',
                      'name_font': {'size': 16, 'bold': True, 'name': 'Arial'},
                      'num_font': {'size': 12, 'italic': False, 'name': 'Microsoft YaHei UI'}})
    chart.set_plotarea({'layout': {'x': 0.1, 'y': 0.1, 'width': 0.8, 'height': 0.8}})
    chart.set_legend({'font': {'name': 'Arial', 'size': 17, 'bold': False},
                      'border': {'color': 'black'},
                      'pattern': {'pattern': 'percent_5', 'fg_color': blue1, 'bg_color': 'white'}
                      })
    # chart.set_style(12)
    i = 0
    for s in series_name_position_list:
        chart.add_series({
            'name': '=' + datasheet_name + '!$B$' + str(s + 1),
            'categories': '=' + datasheet_name + '!$B$' + str(s + 3) + ':$B$' + str(s + 3 + len(list_load) - 1),
            'values': '=' + datasheet_name + '!$F$' + str(s + 3) + ':$F$' + str(s + 3 + len(list_load) - 1),
            'line': {'type': 'automatic', 'color': color_list[i], 'width': 3.0, 'dash_type': 'solid','transparency': 0},
            'marker': {'type': 'circle', 'size': 6
                       # 'fill': {'color': 'red'},
                       # 'border': {'color': 'black'}
                       }
        })
        i += 1
    datasheet.insert_chart(ts_row + 2, 1, chart)
    exlbook.close()

    print ('**************end****************')
    t2 = __time_stamp(logging, 'End')
    sec = t2 - t1
    if sec <= 60:
        print ('------ {}s --------'.format(sec))
    elif 60 < sec <= 3600:
        m = sec // 60
        s = sec % 60
        print ('------ {}m {}s --------'.format(m, s))
