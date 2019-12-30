#----------------modify day 19/11/19
#add mask
import csv
from tkinter import*
from tkinter.filedialog import askopenfilename
from os import listdir
from os.path import isfile, isdir, join
import openpyxl


#-------------------------------------------------
#
#-------------------------------------------------

ItemName=['coreMask','TargetPower','MeasuredPower','PowerDelta','EstimatedPower','PowerIndex','EstimatedPowerAfterTraffic','TSSI','EVM','EVMLimit','HighestEVM','Frequency','Error_ppm','FrequencyError_Hz','Tempsense','PhaseNoise','RampOnTime','RampOffTime','LoLeakage','OFDMA out of left mask[< -30 Mhz]','OFDMA outermost left [-30 to -20 Mhz]','failitems']
WriteItemName=['channel','rate','bandwidth','antenna','MeasuredPower','HighestEVM']
EVMLimit={'1':-5,'2':-5,"5.5":-5,"11":-5,'6':-5,'54':-25,'he0nss1':-5,'he1nss1':-10,'he2nss1':-13,'he3nss1':-16,'he4nss1':-19,'he5nss1':-22,'he6nss1':-25,'he7nss1':-28,'he8nss1':-30,'he9nss1':-32,'he10nss1':-35,'he11nss1':-35}
    #定義寫入起始位置	]

#-------------------------------------------------
#
#-------------------------------------------------
#選擇路徑
def selectPath():   
    path_= askopenfilename()              
    #path_=askdirectory()
    path.set(path_)

#列印路徑
def Print_path():
    
    print(path.get())
    return path.get()

#列印檔案夾內所有檔案    
def ReadFileList():
    mypath=path.get()
    files = listdir(mypath)
    for f in files:
# 產生檔案的絕對路徑
        fullpath = join(mypath, f)
# 判斷 fullpath 是檔案還是目錄
        if isfile(fullpath):
            #print("檔案：", f)
            File_List.append(f)
        elif isdir(fullpath):
            print("目錄：", f)

#-------------------------------------------------
#整理數據
#-------------------------------------------------
def LoadFilePath():
    ReadFileList()
    for i in File_List:
        print(i)


#-------------------------------------------------
#class
#-------------------------------------------------
#設置一個class可用於新增一個字典儲存數據
class ResultValue():
    def __init__(self):
        self.ValueObject={}
    def AddV(self,Vkey,Vvalue):
        self.ValueObject[Vkey]=Vvalue

#load file


#-----------------------------------------
#----- arragement Title-------------------
#-----------------------------------------
#stt='obttxtest dutnumber=0 recordTests="pefl" antenna=1 channel=36 rate=he0nss1 bandwidth=20 sideband=none minPower=10 maxPower=25 stepsize=1 AvgTimes=3 delayBeforeMeasurement=500 pktSize=15000 pktEngineIFS=100 pktcapturetime=1000 fullPktChannelEstimate=off amplitudeTracking=off EVMlimit=-5 freqlimit_high=8 freqlimit_low=-8 high_index=119 low_index=6 powerlimit=2.5 tester=LP nosettle=1 outputfile=%SN_5G_Txtest_HE20HE0_ANT0.csv'
#將字串整理成字典
def AddDrict(stt):
    x=stt.split()     
    oX=ResultValue()                #定義一個class 新增一個字典
    for y in x:
        oY=y.split("=",1)           # 將字串中以=分開形成一個列表 'EVMlimit=-5' ->[EVMlimit][-5]
        if(len(oY)>1):              #刪除只有一個title 資訊 如obttxtest
            oX.AddV(oY[0],oY[1])    #將list [EVMlimit][-5] 形成一個字典EVMlimit:-5
    return oX
#寫入新的excel title
def WriteEXcleTitle(SheetName):
    ItemNumber=0
    SheetColumn=65                     #字串A ASCII 65
    for x in WriteItemName:
        SheetName[chr(SheetColumn)+'1']=x
        #+str(ItemNumber)
        ItemNumber+=1
       #SheetRow+=1
#寫入新的excel data
def WriteEXcleDataColumn(SheetName,StarAddress,DataValue,Endline=0):
    SheetColumn=ord(StarAddress[0])                    #EXCEL column
    SheetRow=int(StarAddress[1])    
    ItemNumber=0
    if(ItemNumber<= Endline if Endline>0 else len(DataValue)):
        for x in DataValue:
            CurretAdress=chr(SheetColumn)+str(SheetRow)
            SheetName[chr(SheetColumn)+str(SheetRow)]=x
            SheetColumn+=1
            ItemNumber+=1
            
#寫入新的excel data           
def WriteEXcleDataRow(SheetName,StarAddress,DataValue):
    SheetColumn=StarAddress[0]    
    SheetRow=int(StarAddress[1])                     #EXCEL row
    ItemNumber=0
    if(ItemNumber<= len(DataValue)):
        for x in DataValue:
            SheetName[chr(SheetColumn)+str(SheetRow)]=x
            SheetRow+=1
            ItemNumber+=1

def WriteEXcleDataRowS(SheetName,StarAddress,DataValue):
    SheetColumn=ord(StarAddress[0])                    #EXCEL column
    SheetRow=int(StarAddress[1:])    
    CurretAdress=chr(SheetColumn)+str(SheetRow)
    #print(CurretAdress)
    SheetName[CurretAdress]=DataValue
    SheetRow+=1
    #print('%%%%%')
    #print(CurretAdress)
    return chr(SheetColumn)+str(SheetRow)
    
def runSort():
    AddressRate='B2'
    AddressCh='A2'
    AddressBw='C2'
    AddressAnt='D2'
    AddressEVM='F2'
    AddressPW='E2'
    CsvData={}
    TitleOfIndex=[]     #存放script
    ScriptOfTitle=[]
    StartOfDataS=[]
    EndOfDataS=[]
    zeroindex=[]
    ObjectList=[]
    File_List=[]
    mypath=Print_path()
    print(mypath)
    exampleFile=open(mypath)
    exampleReader=csv.reader(exampleFile)
    num=0
    #開新EXCEL
    wb=openpyxl.Workbook()
    sheet=wb.active 
    sheet.title='newData'
    NewPath=mypath.replace('.csv','_new.xlsx')
    WriteEXcleTitle(sheet)
    WriteEXcleDataColumn(sheet,'A1',WriteItemName)
    print('8888')
    for row in exampleReader:
        while(row.count('')):
            row.remove('')
        CsvData[num]=row
        num+=1
    i=0
    for num in CsvData:
        if(len(CsvData[num]) == 0):
            zeroindex.append(num)
    for num in zeroindex:
        CsvData.pop(num)
    for num in CsvData:
        if(CsvData[num][0] == 'scriptLine'):
            TitleOfIndex.append(num)
        if(CsvData[num][0] == 'START_OF_DATA_TABLE'):
            StartOfDataS.append(num+1)
        if(CsvData[num][0] == 'TEST_SPECIFICATION'):
            EndOfDataS.append(num-1)
#整理 Script 將整行字串 整理成字典存入list 
    for ScriptOfTitleIndex in TitleOfIndex:
        ScriptOfTitle.append(AddDrict(CsvData[ScriptOfTitleIndex][1] ))    
#for x in ScriptOfTitle:
#    print(x.ValueObject)            
#----------------------------------------
#整理數據成字典存入list
#----------------------------------------
    indexL=0
#取出資料位置
    for x,y in zip(StartOfDataS,EndOfDataS):
    #ObjectList.append(ScriptOfTitle[indexL])
        for ValueIndex in range(x,y):
            indexList=0
            ValueOfObject=ValueIndex
        #定義一個class初始字典
            ValueOfObject=ResultValue()
        #將字典存入             
            ObjectList.append(ValueOfObject)
        #取出數據寫入字典
            for keyL,valueL in zip(ItemName,CsvData[ValueIndex]):
            #若第一次寫入存入 Value Title 
                if(indexList==0):
                    ValueOfObject.AddV(CsvData[TitleOfIndex[indexL]][0],ScriptOfTitle[indexL].ValueObject)
                    indexList+=1
                ValueOfObject.AddV(keyL,valueL)
    #print("-----------------------------") 
    #print(ObjectList[indexL].ValueObject.items())
        indexL+=1
  
    print("-----------------------------")    
#for ListIndex in ObjectList:
    print('ObjectList',len(ObjectList))
    channelS=''
    Cant=''
    Crate=''
    CBW=''
    ratelimit=''
    WorstEVMIndex=0
    WorstEVMIndexList=[]
    EmpIndex=0
    Ochannel=''
    Oant=''
    Orate=''
    OBW=''
    SaveEVM=''
    WriteHiWe=0
    PowerLevel=0
#RateStr=ObjectList[2].ValueObject['scriptLine']
    for y in ObjectList:
        
        RateStr=y.ValueObject['scriptLine']
        OHighestEVM=y.ValueObject['HighestEVM']
        OFDMAMask=y.ValueObject['OFDMA out of left mask[< -30 Mhz]']
        OFDMALeft=y.ValueObject['OFDMA outermost left [-30 to -20 Mhz]']
        Orate=RateStr['rate']
        ratelimit=RateStr['rate']
        Ochannel=RateStr['channel']
        Oant=RateStr['antenna']
        OBW=RateStr['bandwidth']
        CheckValue=''
      
        
        #判斷參數是否為'Highest EVM' is ture  就是第一筆資料 ,Ochannel!=channelS 代表不同筆資料
        if(OHighestEVM !='Highest EVM'):
            if(WriteHiWe==0):
                SaveEVM=WorstEVMIndex
            ShowValue=ObjectList[SaveEVM].ValueObject
            #print(ShowValue['HighestEVM'])
            channelS=Ochannel
            Cant=Oant
            Crate=Orate
            CBW=OBW
            EmpIndex=1
            if(WriteHiWe==0 and float(OHighestEVM)>EVMLimit[ratelimit]):#EVMLimit[ratelimit]
                CheckValue=ObjectList[SaveEVM-1].ValueObject
                print(	'OFDMA out of left mask',OFDMAMask)
                print(	'OFDMA outermost left',OFDMALeft)
                if(CheckValue['HighestEVM'] != 'Highest EVM'):
                    SaveEVM=WorstEVMIndex-1
                WriteHiWe=1
            #print('channel',Ochannel)
            #print('HighestEVM',OHighestEVM)
            #print(type(OHighestEVM))
        #print(Ochannel,channelS)
        #print(Oant,Cant)
        #print(Orate,Crate)
        #print(OBW,CBW)
        if((Ochannel!=channelS or Oant != Cant or Orate!=Crate or OBW != CBW or WorstEVMIndex == len(ObjectList)-1 ) and EmpIndex==1 and SaveEVM != ' '):
            WriteHiWe=0
            WorstEVMIndexList.append(SaveEVM)
            #print('&&&&&&&&&&&')
            #print(SaveEVM)
            #print(Ochannel)
            #print(OHighestEVM)
        #print(WorstEVMIndex)
        WorstEVMIndex+=1
    print("-----------------------------")
    
    ItemNumber=0
    #Cchannel=''
    #Cant=''
    #Crate=''
    #CBW=''
    #Cpower=''
    #Ochannel=''
    #Oant=''
    #Orate=''
    #OBW=''
    #Opower=100
    minPowerIndexList=[]
    PowerIndex=0
    RecFlag=0
#------------------寫入excel--------------------------------------------
    for Oindex in WorstEVMIndexList:
        ShowValue=ObjectList[Oindex].ValueObject
        Oindex=ResultValue()
        Oindex.Channel=ShowValue['scriptLine']['channel']
        Oindex.Rate=ShowValue['scriptLine']['rate']
        Oindex.Bandwidth=ShowValue['scriptLine']['bandwidth']
        Oindex.MPower=ShowValue['MeasuredPower']
        Oindex.Ant=ShowValue['scriptLine']['antenna']

        if((len(minPowerIndexList)<0)):        
            minPowerIndexList.append(Oindex)
        else:
            for CurrentMode in minPowerIndexList:
                if(CurrentMode.Channel == Oindex.Channel and CurrentMode.Rate == Oindex.Rate and CurrentMode.Bandwidth == Oindex.Bandwidth):
                    RecFlag=1
                    if(CurrentMode.MPower>Oindex.MPower):
                        CurrentMode.MPower=Oindex.MPower
                        CurrentMode.Ant=Oindex.Ant
        #若list不存在同一個mode將其存入list
        if(RecFlag==0):
            minPowerIndexList.append(Oindex)
        else:
            RecFlag=0
        
        if(Oindex.Bandwidth=='40'):
            print('===============')
            print(Oindex.Channel)
            print(Oindex.Rate)
            print(Oindex.Bandwidth)
            print(Oindex.MPower)
            print('RecFlag',RecFlag)
            print('===============')     
        ItemNumber+=1
        #print('--------------'+str(ItemNumber))
        #print('rate:',ShowValue['scriptLine']['rate'])
   
        #print('channel:',ShowValue['scriptLine']['channel'])
        #print('bandwidth:',ShowValue['scriptLine']['bandwidth'])
        #print('antenna:',ShowValue['scriptLine']['antenna'])
        #print('HighestEVM:',ShowValue['HighestEVM'])
        #print('MeasuredPower:',ShowValue['MeasuredPower'])
        #print(ShowValue['scriptLine'])
    print('********',ItemNumber)
    for x in minPowerIndexList:
        AddressCh=WriteEXcleDataRowS(sheet,AddressCh,x.Channel)
        AddressRate=WriteEXcleDataRowS(sheet,AddressRate,x.Rate)
        AddressBw=WriteEXcleDataRowS(sheet,AddressBw,x.Bandwidth)
        AddressAnt=WriteEXcleDataRowS(sheet,AddressAnt,x.Ant)
        AddressPW=WriteEXcleDataRowS(sheet,AddressPW,x.MPower)
        #AddressEVM=WriteEXcleDataRowS(sheet,AddressEVM,ShowValue['HighestEVM'])  
        print(x.Channel)
        print(x.Rate)
        print(x.Bandwidth)
        print(x.MPower)
    wb.save(NewPath)
    print(len(minPowerIndexList))    
    print('Finish')
#print(WorstEVMIndexList)
#print(StartOfDataS)
#----------------------------
#write in excle
#----------------------------
#    wb=openpyxl.Workbook()
#    sheet=wb.active 
#    sheet.title='newData'
#    NewPath=mypath.replace('.csv','_new.xlsx')
#    print(NewPath)
#    wb.save(NewPath)    
#------------------------------   
# GUI 
root=Tk()
path=StringVar()

#顯示"目標路徑"
Label(root,text="目標路徑:").grid(row=0,column=0) 
#顯示 選擇路徑
Entry(root,textvariable=path).grid(row=0,column=1)
#選擇路徑按鈕
Button(root,text="路徑選擇",command=selectPath).grid(row=0,column=2)
#確認執行按鈕
Button(root,text="OK",command=runSort).grid(row=0,column=3)
root.mainloop()

        
        



    

