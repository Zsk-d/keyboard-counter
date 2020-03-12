# Author:zsk-d

import pythoncom,PyHook3,os,_thread,json,time
import exportUtil

keyPressMap = None
def keyDownEvent(event):
    if event.Key in keyPressMap: keyPressMap[event.Key] = keyPressMap[event.Key] + 1
    else: keyPressMap[event.Key] = 1
    # print(event.Key)
    return True

def printStartupString(keyPressMap):
    if 'null' in keyPressMap :del keyPressMap['null']
    dataList = sorted(keyPressMap.items(),key=lambda x:x[1],reverse=True)[:10]
    dataString = '└'
    for data in dataList: dataString = dataString + " {}:{},".format(data[0],data[1])
    print("[-----------------< 按键记录 >-----------------]\r\n次数排序：\r\n{}\r\n".format(dataString))

def saveKeyPressRecordThread():
    while True:
        time.sleep(5*60)
        dataString = json.dumps(keyPressMap)
        dataFileName = "data/keyPressRecord.txt"
        dataFile = open(dataFileName,mode='w')
        dataFile.write(dataString)
        printStartupString(keyPressMap) # 输出按键记录信息
        dataFile.close()

def loadDataFile():
    global keyPressMap
    dataFileName = "data/keyPressRecord.txt"
    if not os.path.exists(dataFileName): keyPressMap = {}
    else:
        keyPressMap = eval(open(dataFileName,mode='r',encoding="utf-8").read(os.stat(dataFileName).st_size))
    printStartupString(keyPressMap)

def exportDataFile():
    global keyPressMap
    exportUtil.exportKeyRecordExcel(keyPressMap)

def init():
    keyHookManager = PyHook3.HookManager()
    keyHookManager.KeyDown = keyDownEvent
    # 注册hook并执行
    keyHookManager.HookKeyboard()
    loadDataFile()
    _thread.start_new_thread(saveKeyPressRecordThread,())