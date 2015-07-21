__author__ = 'yalguo'
import xdrlib, sys
import xlrd

# minpower = 9.9995
# maxpower = 20.0005

minpower = 13.0098
maxpower = 23.0108
thredhold = 0.0005

statecol = 0
timecol = 1
SFNcol = 2
framenumcol = 3
RNTIcol = 4
transblocksizecol1 = 13
transblocksizecol2 = 18
layernumcol = 23
CW1AckNackcol = 54
CW2AckNackcol = 55
mappinginfocol = 63

powercol = [6,8,10,12]
mapinfocol = [5,7,9,11]

statelist = []
timelist = []
SFNlist = []
framenumlist = []
RNTIlist = []
framelist = []
#layernumlist = []
AckNacklist = []
mappingCCElist = []
powerlist = []
rowlist = []
errorpowerrowlist = []
errorrowlist = []

scell1timelist = []
scell1framelist = []

scell2timelist = []
scell2framelist = []

checkRNTI = '0x26AF'

def open_excel(file= 'file.xls'):
    try:
        data= xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)


def read_DLSCHDATTX_L1CELLTX(dlschfile='file1.xls', l1cellfile='file2.xls'):
    dlschdata = open_excel(dlschfile)
    dlschtable = dlschdata.sheet_by_index(0)

    dlschnrows = dlschtable.nrows-2
    #dlschncols = dlschtable.ncols

    l1celldata = open_excel(l1cellfile)
    l1celltable = l1celldata.sheet_by_index(0)

    l1cellnrows = l1celltable.nrows-2
    currow = 5

    for rownum in range(6, dlschnrows):
        dlschRNTI = dlschtable.cell(rownum,RNTIcol).value
        if dlschRNTI == checkRNTI:
            dlschtime = dlschtable.cell(rownum,timecol).value
            dlschSFN = dlschtable.cell(rownum,SFNcol).value
            dlschframenum = dlschtable.cell(rownum,framenumcol).value
            dlschmapinfo = dlschtable.cell(rownum,mappinginfocol).value
            dlschlayernum = dlschtable.cell(rownum,layernumcol).value
            dlschcw1acknack = dlschtable.cell(rownum,CW1AckNackcol).value
            dlschcw2acknack = dlschtable.cell(rownum,CW2AckNackcol).value

            l1celltime = l1celltable.cell(currow,timecol).value
            l1cellSFN = l1celltable.cell(currow,SFNcol).value
            l1cellframenum = l1celltable.cell(currow,framenumcol).value
            l1cellframe = l1cellSFN*10+l1cellframenum

            statelist.append(dlschtable.cell(rownum,statecol).value)
            timelist.append(dlschtime)
            SFNlist.append(dlschSFN)
            framenumlist.append(dlschframenum)
            RNTIlist.append(dlschRNTI)
            dlschframe = dlschSFN*10+dlschframenum
            framelist.append(dlschframe)
            mappingCCElist.append(dlschmapinfo)
            rowlist.append(rownum+1)

            #layernumlist.append(dlschlayernum)
            if dlschlayernum == '-':
                AckNacklist.append(dlschcw1acknack)
            elif dlschlayernum == 1:
                dlschtrans1 = dlschtable.cell(rownum,transblocksizecol1).value
                dlschtrans2 = dlschtable.cell(rownum,transblocksizecol2).value
                if dlschtrans2 != 0 and dlschtrans1 != '-':
                    AckNacklist.append(dlschcw1acknack)
                elif dlschtrans2 != 0  and dlschtrans1 != '-':
                    AckNacklist.append(dlschcw2acknack)
                else:
                    AckNacklist.append(3)
            elif dlschlayernum == 2:
                if dlschcw1acknack == 1 and dlschcw2acknack == 1:
                    AckNacklist.append(dlschcw1acknack)
                elif dlschcw1acknack == 0 or dlschcw2acknack == 0:
                    AckNacklist.append(0)
                elif dlschcw1acknack == 2 or dlschcw2acknack == 2:
                    AckNacklist.append(2)
                else:
                    AckNacklist.append(3)
            else:
                assert(0)


            while l1celltime < dlschtime and currow <= l1cellnrows:
                currow = currow+1
                l1celltime = l1celltable.cell(currow,timecol).value
                l1cellSFN = l1celltable.cell(currow,SFNcol).value
                l1cellframenum = l1celltable.cell(currow,framenumcol).value
                l1cellframe = l1cellSFN*10+l1cellframenum


            while l1celltime == dlschtime and l1cellframe != dlschframe and currow <= l1cellnrows:
                currow = currow+1
                l1celltime = l1celltable.cell(currow,timecol).value
                l1cellSFN = l1celltable.cell(currow,SFNcol).value
                l1cellframenum = l1celltable.cell(currow,framenumcol).value
                l1cellframe = l1cellSFN*10+l1cellframenum
            #
            # while l1celltime == dlschtime and l1cellSFN == dlschSFN and \
            #         l1cellframenum < dlschframenum and currow <= l1cellnrows:
            #     currow = currow+1
            #     l1celltime = l1celltable.cell(currow,timecol).value
            #     l1cellSFN = l1celltable.cell(currow,SFNcol).value
            #     l1cellframenum = l1celltable.cell(currow,framenumcol).value

            if l1celltime == dlschtime and l1cellframe == dlschframe:
                addedpowerflag = False
                for i in range(0, 4):
                    l1cellmapinfo = l1celltable.cell(currow,mapinfocol[i]).value
                    if l1cellmapinfo == dlschmapinfo:
                        powerlist.append(l1celltable.cell(currow, powercol[i]).value)
                        addedpowerflag = True
                        break

                if addedpowerflag == False:
                    powerlist.append(0)
            else:
                powerlist.append(0)

def twocc_power_check():
    maxloop = 8
    acknackindex = 0
    powerindex = 1
    previouspower = 0
    currentpower = 0
    needtoupatepower = True

    scell1index = 0
    scell2index = 0
    isscell1sch = False

    while powerindex < len(statelist):
        ackframe = framelist[acknackindex]
        effectpowerframe = ackframe + 8
        if effectpowerframe >= 10240:
            effectpowerframe %= 10240

        if needtoupatepower:
            powerframe = framelist[powerindex]
            previouspower = powerlist[powerindex-1]
            currentpower = powerlist[powerindex]

        while scell1index < len(scell1timelist) and timelist[acknackindex] > scell1timelist[scell1index]:
            scell1index += 1

        while scell1index < len(scell1timelist) and \
                        timelist[acknackindex] == scell1timelist[scell1index] and \
                ((abs(framelist[acknackindex]-scell1framelist[scell1index])<5000 and \
                         framelist[acknackindex] > scell1framelist[scell1index]) or \
                (abs(framelist[acknackindex]-scell1framelist[scell1index])>5000 and \
                             framelist[acknackindex] < scell1framelist[scell1index])):
            scell1index += 1

        if scell1index < len(scell1framelist) and framelist[acknackindex] == scell1framelist[scell1index]:
            isscell1sch = True


        if not isscell1sch and len(scell2timelist) > 0:
            while scell2index < len(scell2timelist) and timelist[acknackindex] > scell2timelist[scell2index]:
                scell2index += 1

            while scell2index < len(scell2timelist) and \
                            timelist[acknackindex] == scell2timelist[scell2index] and \
                    ((abs(framelist[acknackindex]-scell2framelist[scell2index])<5000 and \
                             framelist[acknackindex] > scell2framelist[scell2index]) or \
                    (abs(framelist[acknackindex]-scell2framelist[scell2index])>5000 and \
                                 framelist[acknackindex] < scell2framelist[scell2index])):
                scell2index += 1

            if scell2index < len(scell2framelist) and framelist[acknackindex] == scell2framelist[scell2index]:
                isscell1sch = True

        if effectpowerframe == powerframe:
            if (AckNacklist[acknackindex] == 0 or ((not isscell1sch) and AckNacklist[acknackindex] == 1)) and \
                            (previouspower-0.01) >= minpower:
                previouspower = previouspower - 0.01
                isscell1sch = False
            elif AckNacklist[acknackindex] == 2 and (previouspower + 0.99) <= maxpower:
                previouspower = previouspower + 0.99

            if abs(currentpower-previouspower) > thredhold:
                errorrowlist.append(rowlist[acknackindex])
                errorpowerrowlist.append(rowlist[powerindex])

            acknackindex = acknackindex + 1
            powerindex = powerindex + 1
            needtoupatepower = True

        elif (abs(effectpowerframe-powerframe)<5000 and effectpowerframe > powerframe) or\
                (abs(effectpowerframe-powerframe)>5000 and effectpowerframe < powerframe):
            if abs(currentpower-previouspower) > thredhold:
                errorrowlist.append(rowlist[acknackindex])
                errorpowerrowlist.append(rowlist[powerindex])

            powerindex = powerindex + 1
            needtoupatepower = True

        else:
            if (AckNacklist[acknackindex] == 0 or ((not isscell1sch) and AckNacklist[acknackindex] == 1)) and \
                            (previouspower-0.01) >= minpower:
                previouspower = previouspower - 0.01
                isscell1sch = False
            elif AckNacklist[acknackindex] == 2 and (previouspower + 0.99) <= maxpower:
                previouspower = previouspower + 0.99

            acknackindex = acknackindex + 1
            needtoupatepower = False

def read_DLSCHDATTX_L1CELLTX1(dlschfile='file1.xls'):
    dlschdata = open_excel(dlschfile)
    dlschtable = dlschdata.sheet_by_index(0)

    dlschnrows = dlschtable.nrows-2
    for rownum in range(6, dlschnrows):
        dlschRNTI = dlschtable.cell(rownum,RNTIcol).value
        if dlschRNTI == checkRNTI:
            dlschtime = dlschtable.cell(rownum,timecol).value
            dlschSFN = dlschtable.cell(rownum,SFNcol).value
            dlschframenum = dlschtable.cell(rownum,framenumcol).value

            scell1timelist.append(dlschtime)
            dlschframe = dlschSFN*10+dlschframenum
            scell1framelist.append(dlschframe)

def read_DLSCHDATTX_L1CELLTX2(dlschfile='file1.xls'):
    dlschdata = open_excel(dlschfile)
    dlschtable = dlschdata.sheet_by_index(0)

    dlschnrows = dlschtable.nrows-2

    for rownum in range(6, dlschnrows):
        dlschRNTI = dlschtable.cell(rownum,RNTIcol).value
        if dlschRNTI == checkRNTI:
            dlschtime = dlschtable.cell(rownum,timecol).value
            dlschSFN = dlschtable.cell(rownum,SFNcol).value
            dlschframenum = dlschtable.cell(rownum,framenumcol).value

            scell2timelist.append(dlschtime)
            dlschframe = dlschSFN*10+dlschframenum
            scell2framelist.append(dlschframe)


def main():
    read_DLSCHDATTX_L1CELLTX("D:\\project\\Project_EV\\Log\\LTE2243-A-o_003_02\\DLSCHDATTX_SC01.xls", \
                             "D:\\project\\Project_EV\\Log\\LTE2243-A-o_003_02\\L1CELLTX_SC01.xls")

    read_DLSCHDATTX_L1CELLTX1("D:\\project\\Project_EV\\Log\\LTE2243-A-o_003_02\\DLSCHDATTX_SC13.xls")
    # read_DLSCHDATTX_L1CELLTX2("D:\\project\\Project_EV\\Log\\LTE2243-A-o_004_07\\DLSCHDATTX_SC01.xls")

    twocc_power_check()

    for i in range(0, 3000):
        print rowlist[i], timelist[i], SFNlist[i], framenumlist[i], RNTIlist[i], AckNacklist[i], mappingCCElist[i], \
            powerlist[i], framelist[i]

    for i in range(0, len(errorpowerrowlist)):
        print errorpowerrowlist[i]



main()