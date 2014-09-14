import os,re,subprocess

APPBUKA    = "C:\Program Files\Google\Chrome\Application\chrome.exe"

def saveState(theDir,theStart,theNum):
    with open('MassOpener.txt','w') as f:
        f.write('%s\n%s\n%s\n'%(theDir,theStart,theNum))

def loadState():
    r = list(open('MassOpener.txt','r'))
    r[0] = r[0].rstrip('\n')
    r[1] = int(r[1])
    r[2] = int(r[2])
    return r[0],r[1],r[2]

def openByCMD(theDir,theStart=0,theNum=30):
    windowsCMD   = '"%s" '%(APPBUKA)
    fileNameList = os.listdir(theDir)
    fileCount    = 0
    i            = theStart
    while fileCount <= theNum:
        if re.search('.swf',fileNameList[i]) == None:
            continue
        filePath   = "%s%s"%(theDir,fileNameList[i])
        windowsCMD = windowsCMD+' "'+filePath+'"'
        fileCount  = fileCount + 1
        i          = i + 1
    if fileCount > 0:
        subprocess.call(windowsCMD)

def getInput():
    theDir,theStart,theNum = loadState()

    while True:
        theDirInput = input('Masukkan direktori (atau gunakan %s ): '%(theDir))
        if not theDirInput:
            break
        try:
            os.listdir(theDirInput)
            theDir = theDirInput
            break
        except:
            print('Direktori Salah.')

    while True:
        theStartInput = input('Masukkan urutan awal file (atau gunakan %s ): '%(theStart))
        if not theStartInput:
            break
        try:
            theStart = int(theStartInput)
            break
        except:
            print('Urutan harus integer.')

    while True:
        theNumInput = input('Masukkan berapa banyak file (atau gunakan %s ): '%(theNum))
        if not theNumInput:
            break
        try:
            theNum = int(theNumInput)
            break
        except:
            print('Urutan harus integer.')

    saveState(theDir,theStart,theNum)
    return theDir,theStart,theNum

def massOpen():
    while True:
        theDir,theStart,theNum = getInput()
        openByCMD(theDir,theStart,theNum)
        print('lagi? y/t')
        if input().lower() == 'y':
            continue
        break


massOpen()
