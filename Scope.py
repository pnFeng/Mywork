import win32com.client #imports the pywin32 library
from time import sleep
import visa
#from sub20_gpio import *
from xlsxwriter.workbook import Workbook
exlbook=Workbook("result/ripple_AB_samsung_0x20_0xFF_phase_shedding_full.xlsx")
tempsheet = exlbook.add_worksheet()
scsheet = exlbook.add_worksheet()
ts_row = 0
sc_row=0
testnum=0

Fsw=500e3
Nphase=2.0
Vout=1.1
L=0.68e-6
#assume current ripple at 2A
MaxTimebase=10e-3 #max time base at 0A set as 10ms
MinTimebase=10e-6 #min time base set as 10us
Isweep=[0, 0.001, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.5]
#Isweep=[0.01, 0.1, 0.5]
#Isweep=[0.05, 0.1, 0.2, 0.5]
Vinsweep=[4.25, 10.2, 12, 13.8, 15]
Isweep=Isweep+[x*1 for x in range(1,11)]



rm = visa.ResourceManager()
load = rm.open_resource(u'GPIB0::12::INSTR')
PowerSupply = rm.open_resource('GPIB0::6::INSTR')
meter1 = rm.open_resource('GPIB0::20::INSTR') #Vin
scope=win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1")  #creates instance of the ActiveDSO control
scope.MakeConnection("IP: 157.165.28.196") #Connects to the oscilloscope.  Substitute your IP address

scope.WriteString("*IDN?",1)
print scope.ReadString(100)
sleep(1)
scope.WriteString("""VBS 'app.acquisition.horizontal.maximize="SetMaximumMemory" '""", 1)
scope.WriteString("""VBS 'app.Acquisition.Horizontal.MaxSamples= "50 MS"'""", 1)  # X-timing scale in second
scope.WriteString("""VBS 'app.Acquisition.TriggerMode="auto"'""",1)

#scope.WriteString("""VBS 'app.acquisition.horizontal.horoffsetorigin=0 '""", 1)  # could do horoffset by second also

#step 1 set up the scope. Channel,

###set up the probe channel 1MOhm, 3bits filter, 20MHz
scope.WriteString("""VBS 'app.Acquisition.C1.EnhanceResType="3bits"' """,1)
scope.WriteString("""VBS 'app.Acquisition.C1.BandwidthLimit="20MHz"' """,1)
scope.WriteString("""VBS 'app.Acquisition.C1.Coupling="DC1M"' """,1)#chose between DC50 and DC1M
scope.WriteString("""VBS 'app.Acquisition.C1.LabelsText="Vout"' """,1)

#step 2 set up the measurement in scope
scope.WriteString("""VBS 'app.Measure.P2.ParamEngine="Maximum"'""",1)
scope.WriteString("""VBS 'app.Measure.P2.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P2.View=1'""",1)
scope.WriteString("""VBS 'app.Measure.P3.ParamEngine="Minimum"'""",1)
scope.WriteString("""VBS 'app.Measure.P3.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P3.View=1'""",1)
scope.WriteString("""VBS 'app.Measure.P4.ParamEngine="Mean"'""",1)
scope.WriteString("""VBS 'app.Measure.P4.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P4.View=1'""",1)
scope.WriteString("""VBS 'app.Measure.P5.ParamEngine="Frequency"'""",1)
scope.WriteString("""VBS 'app.Measure.P5.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P5.View=1'""",1)
scope.WriteString("""VBS 'app.Measure.P6.ParamEngine="Period"'""",1)
scope.WriteString("""VBS 'app.Measure.P6.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P6.View=1'""",1)
scope.WriteString("""VBS 'app.Measure.P1.ParamEngine="PeakToPeak"'""",1)
scope.WriteString("""VBS 'app.Measure.P1.Source1="C1"'""",1)
scope.WriteString("""VBS 'app.Measure.P1.View=1'""",1)
scope.WriteString("VBS app.Measure.ShowMeasure=1",1)
scope.WriteString("VBS app.Measure.StatsOn=1",1)
#step 3 set up the time scale and voltage scale
#at 0A time scale 10ms/div
#at <2A*Nphase, Period time = NPhase*Ipeak/(2*Iout*Fsw), Ipeak=(Vin-Vout)*Vout/(Vin*Fsw*L) (assume 2-3A peak current is also fine), Timebase=2*Period
def screen_capture(Itarget, Vintarget):
    global testnum
    global sc_row

    scsheet.write(sc_row, 0, 'Test#')
    scsheet.write(sc_row, 1, testnum)
    scsheet.write(sc_row, 2, 'Iout')
    scsheet.write(sc_row, 3, Itarget)
    scsheet.write(sc_row, 4, 'Vin')
    scsheet.write(sc_row, 5, Vintarget)

    sc_row +=1
    #screen shot
    scsheet.set_row(sc_row, 300)
    scope.StoreHardcopyToFile("PNG", "", "D:\python_program\\auto_test\\temp_pic\TIFFImage" + str(testnum) + ".PNG")
    scsheet.insert_image('A' + str(sc_row + 1),
                           "D:\python_program\\auto_test\\temp_pic\TIFFImage" + str(testnum) + ".PNG",
                           {'x_scale': 0.5, 'y_scale': 0.5})
    sc_row += 1
    # Capture delay signals

def measurement(Iout, Vin, Vout):
    global ts_row
    global testnum
    Timebase=MaxTimebase
    Ipeak = (Vin - Vout) * Vout / (Vin * Fsw * L)
    #print Ipeak
    if Iout==0:
        Timebase = MaxTimebase
    elif Iout<Ipeak*Nphase: #DCM

        Timebase=2*(Nphase*Ipeak/(Iout*Fsw*2.0))
        #print Timebase/2
    else:                   #CCM 4 period for each division
        Timebase = 4 / Fsw
    #clamp Timebase between MaxTimebase and MinTimebase
    if Timebase>MaxTimebase:
        Timebase=MaxTimebase
    elif Timebase<MinTimebase:
        Timebase = MinTimebase
    #print "VBS app.acquisition.horizontal.horScale=" + str(Timebase)
    scope.WriteString("VBS app.Acquisition.Horizontal.HorScale=" + str(Timebase), 1)

    #voltage level, set at 10mV/div, Vout+40mV top, Vout-30mV bottom. add the trim function in case voltage out of range.
    #step 4 find level at every step before DCM (2*Nphase)
    sleep(1)
    scope.WriteString("VBS 'app.Acquisition.C1.FindScale'",1)#scope auto scale and offset
    sleep(3)
    scope.WriteString("VBS 'app.clearsweeps'",1)
    sleep(1)


    #wait for trigger count over 100 or 50s
    dvalue=0
    maxwaitcnt=10
    #5/31 added, zoom out if Vout is too large
    sleep(2)
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P1").max.result.value'""", 1)
    try :
        Vpk = float(scope.ReadString(80))
    except:
        scope.WriteString("""VBS? 'return=app.Acquisition.C1.VerScale.value'""", 1)
        verscale = float(scope.ReadString(80))
        scope.WriteString("VBS app.Acquisition.C1.VerScale.value=" + str(2 * verscale), 1)
        sleep(2)
        scope.WriteString("""VBS? 'return=app.Measure.Measure("P1").max.result.value'""", 1)
        Vpk = float(scope.ReadString(80))
    #readback channel resolution mV/div
    scope.WriteString("""VBS? 'return=app.Acquisition.C1.VerScale.value'""", 1)
    verscale = float(scope.ReadString(80))
    #print verscale, type(verscale)
    total_room=8*verscale
    if Vpk>0.6*total_room:
        scope.WriteString("VBS app.Acquisition.C1.VerScale.value=" + str(2*verscale), 1)
        #double the room
    #end of 5/31 edit
    while dvalue<30 and maxwaitcnt>0:
        sleep(10)
        scope.WriteString("""VBS? 'return=app.Measure.Measure("P1").num.result.value'""", 1)
        dvalue = int(scope.ReadString(80))
        maxwaitcnt=maxwaitcnt-1
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P1").max.result.value'""", 1)
    Vpk_mean=float(scope.ReadString(80))
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P2").max.result.value'""", 1)
    Vout_max=float(scope.ReadString(80))
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P3").min.result.value'""", 1)
    Vout_min=float(scope.ReadString(80))
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P4").mean.result.value'""", 1)
    Vout_mean=float(scope.ReadString(80))
    scope.WriteString("""VBS? 'return=app.Measure.Measure("P6").mean.result.value'""", 1)
    Period=float(scope.ReadString(80))
    print Iout,Vpk_mean,Vout_max,Vout_min,Vout_mean,Period
    tempsheet.write(ts_row, 0, Iout)
    tempsheet.write(ts_row, 1, Vpk_mean*1000)
    tempsheet.write(ts_row, 2, Vout_max)
    tempsheet.write(ts_row, 3, Vout_min)
    tempsheet.write(ts_row, 4, Vout_mean)
    tempsheet.write(ts_row, 5, testnum)
    ts_row=ts_row+1



load.write('CHANNEL 1;')
load.write('LOAD ON')
for Vintarget in Vinsweep:
    PowerSupply.write('VOLT ' + str(Vintarget) + '')
    #set up worksheet
    tempsheet.write(ts_row, 0, 'Vin = '+str(Vintarget)+"V")
    ts_row = ts_row + 1
    tempsheet.write(ts_row, 0, 'Iout (A)')
    tempsheet.write(ts_row, 1, 'Vripple (mV)')
    tempsheet.write(ts_row, 2, 'Vout max (V)')
    tempsheet.write(ts_row, 3, 'Vout min (V)')
    tempsheet.write(ts_row, 4, 'Vout mean (V)')
    tempsheet.write(ts_row, 5, 'Test number')
    ts_row=ts_row+1
    load.write('LOAD ON')
    Vin_comp = Vintarget
    for Itarget in Isweep:
        if Itarget<1:
            load.write(':MODE CCL;')
        else:
            load.write(':MODE CCH;')
        load.write('CURR:STAT:L1 ' + str(Itarget) + ';')
        # Vin Compensation

        daq_comp = abs(float(meter1.query('MEAS:VOLT:DC?')))#add abs in case wrong polarity
        if abs(Vintarget - daq_comp)>0.005:
            daq_add = Vintarget - daq_comp
            Vin_comp = Vin_comp + daq_add
            # Protection just in case it goes waaaay over Vin limit
        if Vin_comp > Vintarget+1:
            PowerSupply.write('VOLT ' + str(Vintarget) + '')
            Vin_comp=Vintarget
        elif Vin_comp<Vintarget-1:
            PowerSupply.write('VOLT ' + str(Vintarget) + '')
            Vin_comp=Vintarget
        else:
            PowerSupply.write('VOLT ' + str(Vin_comp) + '')
        sleep(1)
        measurement(Itarget, Vintarget, Vout)
        screen_capture(Itarget, Vintarget)
        testnum = testnum + 1
    load.write('LOAD OFF')
exlbook.close()
load.write('LOAD OFF')
PowerSupply.write('OUTP OFF')
scope.Disconnect()