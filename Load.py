from Mylib.bench_resource import brsc
import time

# el1 = brsc.Eload('USB0::0x0A69::0x083E::636002002010::0::INSTR')
# el1.cc(il=0.5, ch=2)
# el1.cc(il=1.5, ch=3)
# el1.dc(0.5, 0, 1)
# el1.on([1, 2, 3])
# time.sleep(10)

# p1 = brsc.DPS('USB0::0x2A8D::0x1202::MY58270817::0::INSTR')
# p1.on(5, 2, 1)
#
# for i in [0, 0.25, 0.5, 0.75, 1, 1.5, 2]:
#     el1.cc(i, 1)
#     el1.on([1])
#     time.sleep(1)
#     c = p1.meas(2)
#     cout = el1.meas()
#     print(c, cout)

brsc.Scope('USB0::0x05FF::0x1023::3808N60392::INSTR')
