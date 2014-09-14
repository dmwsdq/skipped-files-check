'''
skippedFilesCheck.py adalah script kecil untuk mendeteksi adanya file
terlewatkan selama pengumpulan file dalam jumlah banyak, selama file tersebut
memiliki nama file dengan pola tertentu. Pola dimaksud adalah:

SDOC | DOCNUM | EDOC | SPAGE | PAGENUM | EPAGE

Contoh: "doc=M1&page=55.swf"

SDOC dan EDOC merupakan string di awal dan akhir DOCNUM (S = start, E = end),
begitu pula dengan SPAGE dan EPAGE terhadap PAGENUM. SDOC, EDOC, SPAGE, EPAGE
di sini di sebut SEPARATOR.
DOCNUM adalah suatu nama dokumen yang berurutan, memiliki halaman-halaman
yang urutannya ditulis sebagai PAGENUM. Misal dokumen M1 halaman 55, dst.
Contoh di atas apabila dipecah berurutan dari SDOC hingga EPAGE adalah sbb:

"doc=" | "M1" | "&" | "page=" | "55" | ".swf"

Cara kerjanya adalah dengan melihat unsur DOCNUM dan PAGENUM, yang menjadi
unsur pembeda suatu file dengan file lainnya. Berkaitan dengan unsur pembeda
tersebut, ada 2 lingkup:
1. File-file se-direktori, meliputi seluruh doc dan page.
Doc belum tentu menunjukkan suatu urutan. Apabila menunjukkan urutan, dapat digunakan
fungsi getMissingDocList() untuk mengecek doc terlewat berdasarkan urutan tersebut.
2. File-file se-doc, meliputi seluruh page dalam satu doc.
Page pasti menunjukkan suatu urutan.
3. Satu file, menunjukkan satu page.
Informasi doc dan page disimpan dalam nama file.

Penerapan pada unduh SWF dari cache dan konversi SWF ke PDF:
- Mengecek kelengkapan file SWF terunduh (fase 1)
- Mengecek kelengkapan file PDF terkonversi (fase 2)
- Secara otomatis membuka 30 file SWF yang terlewatkan
'''

import re,subprocess,time,sys,os

'''Deklarasikan konstanta di sini. 1 merujuk pada pengecekan SWF, 2 pengecekan PDF.
Format SEPARATOR: [SDOC, EDOC, SPAGE, EPAGE]
SEPDOCNUM = Separator [AWAL,AKHIR] hanya diperlukan jika  doc menunjukkan urutan dan fungsi getMissingDocList
akan digunakan. 
UODOC = [List of Unordered Doc] hanya diperlukan jika getMissingDocList akan digunakan,
untuk dikecualikan dari pengecekan urutan.
LIMITBY = Berapa file maksimum yang boleh dibuka jika akan menggunakan openByCMD.'''
APPBUKA    = "C:\Program Files\Google\Chrome\Application\chrome.exe"
DIR1       = "D:\Dokumen\Buku\Memilikiformatswf\\"
DIR2       = "D:\Dokumen\Buku\Hasilkonversi\\"
SEPARATOR1 = ['doc=','&','page=','.swf']
SEPARATOR2 = ['doc=','&','page=','.swf.pdf']
SEPDOCNUM  = ['M','']
UODOC      = ['TOC']
LIMITBY    = 30

def isLastNameMatch(theFileName,theSeparator):
    '''Menghasilkan True jika unsur terakhir (ekstensi file) dari theFileName = theSeparator[3].
    Digunakan oleh: getDocPage()'''
    r = type(re.search('(.*)%s$' % (theSeparator[3]),  theFileName)) != type(None)
    return r

def getDocPage(theFileName,theSeparator):
    '''Menghasilkan tuple (theDoc,thePage) dari theFileName berekstensi theSeparator[3].
    Jika theFileName bukan berekstensi theSeparator[3], maka menghasilkan tuple ('OTHER',theFileName).
    Digunakan oleh: getDocPageTable(), Menggunakan: isLastNameMatch()'''
    if isLastNameMatch(theFileName,theSeparator):
        theDoc  = re.search('%s(.*)%s' % (theSeparator[0], theSeparator[1]), theFileName).group(1)
        thePage = re.search('%s(.*)%s' % (theSeparator[2], theSeparator[3]), theFileName).group(1)
        return theDoc,int(thePage)
    else:
        return 'OTHER',theFileName

def getDocPageTable(theDir,theSeparator):
    '''Menghasilkan docPageTable dari seluruh file di bawah theDir. docPageTable merupakan tabel yang
    memuat seluruh file di bawah direktori theDir, disusun dengan dikelompokkan berdasarkan doc. Formatnya
    adalah sbb:
    docPageTable = {'M1':[1,2,3,...,55],'M2':[1,2,3,...,48],...,'OTHER':['anu.py','biji.rar',...]}
    Digunakan oleh: checkSkippedFiles(), Menggunakan: getDocPage()'''
    docPageTable = {}
    fileNameList = os.listdir(theDir)
    for theFileName in fileNameList:
        theDoc,thePage = getDocPage(theFileName,theSeparator)
        if theDoc not in docPageTable:
            docPageTable[theDoc] = []
        docPageTable[theDoc].append(thePage)
    return docPageTable

def getMissingPageTable(docPageTable):
    '''Menghasilkan missingPageTable dari docPageTable. missingPageTable merupakan tabel yang
    memuat dugaan page yang dilewatkan, jika dilihat dari urutan, dalam setiap doc di docPageTable.
    Dengan demikian missingPageTable tidak memuat file yang tergolong 'OTHER'.
    missingPageTable juga disusun dengan dikelompokkan berdasarkan doc. Formatnya sama dengan docPageTable, sbb:
    missingPageTable = {'M1':[6,11,40],'M3':[6],...}
    Digunakan oleh: checkSkippedFiles()'''
    missingPageTable = {}
    #lingkup se-direktori
    for theDoc in docPageTable:
        #lingkup se-doc
        if theDoc == 'OTHER':
            continue
        theDocPages = docPageTable[theDoc]
        maxPage     = max(theDocPages)
        for i in range(1,maxPage):
            #lingkup satu file / satu page
            if i not in theDocPages:
                if theDoc not in missingPageTable:
                    missingPageTable[theDoc] = []
                missingPageTable[theDoc].append(i)
    return missingPageTable

def getMissingDocList(docPageTable):
    '''Menghasilkan missingDocList dari docPageTable. missingDocList merupakan daftar yang memuat
    dugaan doc yang dilewatkan, jika dilihat dari urutan doc dalam satu direktori. Fungsi ini hanya dapat
    dipakai jika doc menunjukkan suatu urutan.
    Contoh: missingDocList = ["M5","M8"]
    Digunakan oleh: checkSkippedFiles()'''

    missingDocList = []
    docNumArray    = []

    for theDoc in docPageTable:
        '''Menentukan doc terakhir'''
        if theDoc == 'OTHER':
            continue
        if theDoc in UODOC:
            continue
        docNum = int(re.search('%s(.*)%s' % (SEPDOCNUM[0], SEPDOCNUM[1]), theDoc).group(1))
        docNumArray.append(docNum)
    if len(docNumArray) > 0:
        maxDocNum = max(docNumArray)
    else:
        maxDocNum = 0

    for i in range(1,maxDocNum):
        '''Menyimpan doc terlewatkan dalam missingDocList'''
        if i not in docNumArray:
            missingDocList.append(i)        
    return missingDocList

def displayMissingPageTable(missingPageTable):
    '''Menampilkan semua dugaan page yang terlewatkan.
    Digunakan oleh: checkSkippedFiles()'''
    if len(missingPageTable) > 0:
        for theDoc in missingPageTable:
            print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
            print('!!!!!!!!!!!!!!!!! Halaman terlewat pada doc %s : '%theDoc,end='')
            for thePage in missingPageTable[theDoc]:
                print('%s, '%thePage,end='')
            print('')
    else:
        print('Page lengkap')
    
def displayMissingDocList(missingDocList):
    '''Menampilkan semua dugaan doc yang terlewatkan.
    Digunakan oleh: checkSkippedFiles()'''
    if len(missingDocList) > 0:
        print('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
        print('!!!!!!!!!!!!!!!!! Doc yang terlewat: ',end='')
        for theDoc in missingDocList:
            print('%s, '%theDoc,end='')
        print()
    else:
        print('Doc lengkap')
    
def displayGeneralStatistics(theDir,theSeparator,docPageTable):
    '''Menampilkan statistik secara umum.
    Digunakan oleh: checkSkippedFiles()'''

    print('_____________________________________________________________________________')
    print('Direktori: %s, Separator: %s '%(theDir,theSeparator))
    if 'OTHER' in docPageTable:
        printThis = docPageTable['OTHER']
    else:
        printThis = 'tidak ada'
    print('File selain berekstensi %s : %s '%(theSeparator[3],printThis))

    numOfFile = 0
    for theDoc in docPageTable:
        if theDoc != 'OTHER':
            numOfFile += len(docPageTable[theDoc])
            print('Doc %s | %s h | akhir h.%s'%(theDoc,len(docPageTable[theDoc]),max(docPageTable[theDoc])))
    print('Jumlah semua file yang berekstensi %s: %s '%(theSeparator[3],numOfFile))

def openByCMD(missingPageTable,CMDDir,theSeparator,limitBy):
    '''Membuka hingga sebanyak limitBy file yang terlewatkan (yang ada dalam missingPageTable),
    menggunakan APPBUKA'''
    windowsCMD = '"%s" '%(APPBUKA)
    fileCount  = 0
    for theDoc in missingPageTable:
        for thePage in missingPageTable[theDoc]:
            filePath = "%s%s%s%s%s%s%s"%(CMDDir,theSeparator[0],theDoc,theSeparator[1],theSeparator[2],thePage,theSeparator[3])
            windowsCMD = windowsCMD+' "'+filePath+'"'
            fileCount  = fileCount + 1
            if fileCount >= limitBy:
                break
        if fileCount >= limitBy:
            break
    if fileCount > 0:
        subprocess.call(windowsCMD)

def getMissingPageTableByDir(docPageTable1,docPageTable2):
    '''Membandingkan docPageTable fase 1 dengan fase 2, patokan adalah fase 1.'''
    missingPageTable = {}
    print('____________________________________________________________________')
    for theDoc in docPageTable1:
        if theDoc == 'OTHER':
            continue
        if theDoc not in docPageTable2:
            missingPageTable[theDoc] = []
            for i in range(1,max(docPageTable1[theDoc])):
                missingPageTable[theDoc].append(i)
            continue
        for thePage in docPageTable1[theDoc]:
            if thePage not in docPageTable2[theDoc]:
                if theDoc not in missingPageTable:
                    missingPageTable[theDoc] = []
                missingPageTable[theDoc].append(thePage)
    return missingPageTable

def checkSkippedFiles(theDir,theSeparator,useCMD=False,CMDDir='',CMDSeparator=[],checkMissingDocList=True):
    '''Membungkus fungsi-fungsi di atas untuk digunakan oleh fase 1.'''

    docPageTable     = getDocPageTable(theDir,theSeparator)
    displayGeneralStatistics(theDir,theSeparator,docPageTable)

    missingPageTable = getMissingPageTable(docPageTable)
    displayMissingPageTable(missingPageTable)

    if checkMissingDocList:
        missingDocList   = getMissingDocList(docPageTable)
        displayMissingDocList(missingDocList)
        
    if useCMD:
        openByCMD(missingPageTable,CMDDir,CMDSeparator,LIMITBY)

def compareDir(theDir1,theSeparator1,theDir2,theSeparator2,useCMD=False,CMDDir='',CMDSeparator=[],checkMissingDocList=True):
    '''Membungkus fungsi-fungsi di atas untuk digunakan oleh fase 2.'''

    docPageTable1     = getDocPageTable(theDir1,theSeparator1)
    docPageTable2     = getDocPageTable(theDir2,theSeparator2)
    displayGeneralStatistics(theDir2,theSeparator2,docPageTable2)

    missingPageTable = getMissingPageTableByDir(docPageTable1,docPageTable2)
    displayMissingPageTable(missingPageTable)

    if checkMissingDocList:
        missingDocList   = getMissingDocList(docPageTable2)
        displayMissingDocList(missingDocList)
        
    if useCMD:
        openByCMD(missingPageTable,CMDDir,CMDSeparator,LIMITBY)

    
def initialization():
    '''Membungkus
    checkSkippedFiles() bagi fase 1, dan
    compareDir() bagi fase 2.'''
    #checkSkippedFiles(DIR1,SEPARATOR1,False,  '',[],True)
    compareDir(DIR1,SEPARATOR1,DIR2,SEPARATOR2,True,DIR1,SEPARATOR1,True)

def termination():
    '''Keluar dari script'''
    print("Keluar dalam 60 detik...")
    time.sleep(60)
    sys.exit()


initialization()
termination()
