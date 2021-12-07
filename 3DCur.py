import time
from openpyxl import Workbook
from Mylib.bench_resource import brsc
import argparse
import pkg_resources.py2_warn

path = 'D:\DDR_PMIC\Pinehurst\TDLoadCurrent\data.xlsx'


def scopefun(ts):
    s1 = brsc.Scope('USB0::0x05FF::0x1023::3808N60392::INSTR')
    # s1.scope.write('vbs app.acquisition.Horizontal.HorScale =' + str(ts))
    # s1.scope.write('vbs app.measure.clearsweeps')
    # s1.scope.write('vbs app.acquisition.C1.LabelsText="Vout"')

    # s1.scope.write('vbs app.acquisition.triggermode="auto"')
    # s1.scope.write('vbs app.acquisition.trigger.')
    # s1.scope.write('vbs app.acquisition.trigger.edge.FindLevel')

    # time.sleep(3)
    # c1 = float(s1.scope.query('vbs? return = app.measure.p1.max.result.value'))
    # c2 = float(s1.scope.query('vbs? return = app.measure.p2.min.result.value'))
    # c3 = float(s1.scope.query('vbs? return = app.measure.p1.max.result.value'))
    # c4 = float(s1.scope.query('vbs? return = app.measure.p2.min.result.value'))
    # s1.scope.write('HCSU DEV, JPEG, AREA, DSOWINDOW, PORT, NET')
    # s1.scope.write('SCDP')
    # screen = s1.scope.read_raw()
    # f = open("D:\DDR_PMIC\Pinehurst\TDLoadCurrent\Capture.jpg", 'wb+')
    # f.write(screen)
    # f.close()
    # print(c1, c2, c3, c4)
    # return c1, c2, c3, c4
    pass


# parser = argparse.ArgumentParser()
# parser.add_argument('timescale', help="timescale", type=float)
# args = parser.parse_args()

wb = Workbook()
ws1 = wb.active
scopefun(5e-6)
wb.save(path)
