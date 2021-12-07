# -*- coding: utf-8 -*-

import logging
import time
import pandas as pd
from datetime import datetime
import itertools

from dolphin import *
import ICs
from ICs import Ddr5Device
from tests.test_wrapper import SpyglassHillsTestWrapper

start = time.time()

logname = datetime.now().strftime("%Y%m%d-%H_%M_%S")
logger = logging.getLogger('mylogger')
handle = logging.FileHandler('I2C_filter_fsw_test_{}.log'.format(logname))
logger.addHandler(handle)


def write_read_test(dut):
    write_mode = [
        (0, 1), (0, 2), (0, 3), (0, 4),
        (1, 1), (1, 2), (1, 3),
        (2, 1), (2, 2),
        (3, 1)
    ]

    tRegs = dut.tp_write_read_regs_flat['regs']
    tData = dut.tp_write_read_regs_flat['data']

    tIdx = 0
    for offset, blksize in write_mode:
        # Update the register with the request offset and blksize
        tIdx = (tIdx + 1 ) % 2
        exp_data = tData[tIdx][offset: offset + blksize]
        ret = dut.i2c_write_reg(tRegs[offset], exp_data)
        if ret != 0:
            return "WriteNacked:{}".format(ret)

        # Check the data is updated as expected
        try:
            ret, data = dut.i2c_read_reg(tRegs[offset], blksize)
            if ret != 0:
                return "ReadNacked:{}".format(ret)
            else:
                if isinstance(data, int):
                    data = (data,)
                if data != tuple(exp_data):
                    return "DataError:Err:{}, Exp:{}".format(
                        Ddr5Device.hexListStr(data),
                        Ddr5Device.hexListStr(exp_data))
        except Exception as e:
            return str(e)

        return None


dut = SpyglassHillsTestWrapper(hid=7)
assert (dut.connect(option=OUTPUT_OPTION_3))
dut.set_voltage(1.0)
dut.set_i3c_bitrate(1000, 100)
dut.bus_reset_ext()
dut.util_unlock_idt_registers()

# dut.i2c_write_reg(0xAD, 0xA1)
# dut.i2c_write_reg(0xD4, 0X55)
# ret, RxAD = dut.i2c_read_reg(0xAD)
# assert ret == 0
# ret, RxD4 = dut.i2c_read_reg(0xD4)
# assert ret == 0
# print("RxAD={:02X}, RxD4={:02X}".format(RxAD, RxD4))

dut.i2c_initialize(bEnableVR=False)
time.sleep(0.5)


rw_repeat = 2000
max_err = 100

scan_i2c_freq = [100, 400, 600, 800, 1000]
scan_volt = [1.0, 1.5, 1.8, 2.5, 3.3]

# scan_filters = [ dut.i2c_filter.filters]
scan_Fsw = [((0x29, 0x88), (0x2A, 0x89)),    # Fsw =500KHZ/500KHZ/500KHZ/750KHZ
            ((0x29, 0x99), (0x2A, 0x99))]    # Fsw =750KHZ/750KHZ/750KHZ/750KHZ

record_dict = {}
for Fsw_setting in scan_Fsw:
    dut.set_i3c_bitrate(1000, 100)
    print("-" * 10)
    fsw_key = ""
    for reg, setting in Fsw_setting:
        # filter.i2c_set()
        dut.i2c_write_reg(reg, setting)
        fsw_key += "{:02X}:{:02X};".format(reg, setting)
    record_dict[fsw_key] = {}

    dut.enable_vr(True, "I2C")
    time.sleep(0.5)

    for volt, freq in itertools.product(scan_volt, scan_i2c_freq):
        cur_key = "{:.01f}V / {}".format(volt, freq)
        dut.set_voltage(volt)
        dut.set_i3c_bitrate(1000, freq)

        print("Fsw:{}, Volt: {:.01f}, Freq:{} ".format(fsw_key, volt, freq))
        logger.warning("Fsw:{}, Volt: {:.01f}, Freq:{} ".format(fsw_key, volt, freq))

        errCount = 0
        try:
            for cycle in range(rw_repeat):
                err = write_read_test(dut)
                if err is not None:
                    errCount += 1
                    logger.warning(">>>> Failed at cycle{}: {}".format(cycle, err))
                    print(">>>> Failed at cycle{}: {}".format(cycle, err))
                if errCount >= max_err:
                    assert 0, "Stop on Max error"
            logger.warning(">>>> Passed")
        except:
            dut.free_bus()
        finally:
            record_dict[fsw_key][cur_key] = errCount
            dut.set_i3c_bitrate(1000, 100)

    dut.enable_vr(False, "I2C")
    time.sleep(0.5)

df = pd.DataFrame.from_dict(record_dict)
df.T.to_excel('I2C_filter_fsw_test_{}.xlsx'.format(logname))

print("TimeUsage: {:.1f} Minutes".format((time.time()-start)/60.0))