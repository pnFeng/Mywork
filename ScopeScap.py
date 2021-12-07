from Mylib.bench_resource import brsc
import time
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('freq', help="Load Freq", type=str)
args = parser.parse_args()


def Capture(freq):
    s1 = brsc.Scope('USB0::0x05FF::0x1023::3808N60392::INSTR')
    print(freq)

    '''
    freq = float(freq)
    if freq < 1000:
        N = 2
    else:
        N = 5
    T = 0.1 * N / freq
    # s1.scope.write('vbs app.Acquisition.C2.VerOffset = {}'.format(0))
    s1.scope.write('vbs app.Acquisition.Horizontal.HorScale = {}'.format(T))
    s1.scope.write('vbs app.Acquisition.Horizontal.HorOffset = {}'.format(-3 * T))
    # s1.scope.write('vbs app.Acquisition.Trigger.Source = "C8"')
    # s1.scope.write('vbs app.acquisition.trigger.edge.level = {}'.format(3))
    s1.scope.write('vbs app.acquisition.triggermode = "Signle"')  # Auto Single Normal Stopped
    time.sleep(1)
    s1.SCap('D:\Profile\Desktop\Test\R_' + str(freq) + '.png')
    '''


Capture(args.freq)
