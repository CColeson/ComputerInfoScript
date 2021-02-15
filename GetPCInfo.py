'''need computer name, location/ associate, service tag, desktop type,
   manufacturer, current operating system, last 5 win prod key,
   CPU name, amount of memory, ssd v hdd, ram dimms'''
import os
import subprocess
import shutil
import csv


def usewmic(command):
    out = subprocess.run(command,capture_output=True).stdout.strip().decode('utf-8').split()[1:]
    tmp=''
    for i in out:
        tmp += i + ' '
    return tmp

def tagname(cpuInfo):
    stag = input('\nEnter Service Tag. Enter "False" if pre-built\n').lower()
    compName = stag
    if stag == 'false' or stag == 'f':
        compName = cpuInfo
    return stag.upper(),compName.upper()

def associate():
    return input('\nEnter Associate/ Location\n')

def getType():
    #desktop / aio / laptop
    types = ['Desktop','Laptop','AIO 23"','AIO 27"','Other']
    print("\nWhat type of computer is this?")
    for idx,i in enumerate(types):
        print(''+str(idx)+'.)'+i)
    compType = types[int(input())]
    if compType == types[-1]:
        compType = input('Enter the type of computer\n')
        
    return compType

def getStorage():
    #storage
    total,used,free = shutil.disk_usage("C:")
    return str(total // (2**30)) + ' GB '

def getRam():
    #ram
    mem = subprocess.run('wmic memorychip get capacity',capture_output=True).stdout.strip().decode('utf-8').split()
    return str(int(mem[-1]) // (2**30)) + ' GB'

def findOffice():
    'cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus'
    #Office 2016/2019 (32-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /dstatus'
    #Office 2016/2019 (64-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus'
    #Office 2013 (32-bit) on a 32-bit version of Windows
    'cscript "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" /dstatus'
    #Office 2013 (32-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS" /dstatus'
    #Office 2013 (64-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" /dstatus'
    #Office 2010 (32-bit) on a 32-bit version of Windows
    'cscript "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" /dstatus'
    #Office 2010 (32-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS" /dstatus'
    #Office 2010 (64-bit) on a 64-bit version of Windows
    'cscript "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" /dstatus'
    return ''

def findWinKey():
    return ''
    
def dimms():
    #number of dimms available
    dimmNum = subprocess.run('wmic MEMPHYSICAL get memorydevices',capture_output=True).stdout.strip().decode('utf-8').split()[-1]
    #number of dimms in use
    numSticks = subprocess.run('wmic MEMORYCHIP get capacity',capture_output=True).stdout.strip().decode('utf-8').split()[1:]
    numS = []
    for i in numSticks:
        numS.append(str(int(i) // (2**30)))
    numD = {}
    for i in numS:
        if i in numD:
            numD[i] += 1
        else:
            numD[i] = 1
    string = ''
    for i in numD:
        string += str(numD[i]) + ' x ' + i + ' GB '
    return string


def driveType():
    options = ['HDD','SSD','Other']
    drives = subprocess.run('wmic diskdrive get caption',capture_output=True).stdout.strip().decode('utf-8').split('\n')[1:]
    print('\nThe following drives have been found:')
    for i in drives:
        print(i)
    print('\nUsing this information, what is the C Drive')
    for idx,i in enumerate(options):
        print(''+str(idx)+'.)'+i)
    typ = options[int(input())]
    if typ == options[-1]:
        typ = input('Enter the type of drive\n')
    return typ
        
    
     
    


cpu = usewmic('wmic cpu get name')
assLoc = associate()
stag,name = tagname(cpu)
compType = getType()
manu = usewmic('wmic computersystem get manufacturer')
model = usewmic('wmic computersystem get model')
OS = usewmic('wmic os get caption')
winProd = ''#TODO: make these correct
offProd = ''
ram = getRam()
storage = getStorage() + driveType()
dimmInfo = dimms()



    

with open('computerinfo.csv','a',newline='') as csvfile:
    w = csv.writer(csvfile, delimiter=',', quotechar='"',quoting=csv.QUOTE_MINIMAL)
    w.writerow([name,assLoc,stag,compType,manu,model,OS,winProd,offProd,cpu,ram,storage,dimmInfo])

