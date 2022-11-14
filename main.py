import json
import pandas as pd

class jsonConvertor:

    def __init__(self):
        conf = self.loadConfig()
        if conf != False:
            res,keyAry= self.parsingJsonFile(conf['jsonFile'],conf['inputType'])
            
            if conf['convertToExcel']:
                self.exportToExcel(res,keyAry)
            if conf['convertToJson']:
                self.exportToJson(res)
            input()
        return

    def loadConfig(self):
        config = {}
        print('Start Load Config')
        with open('config.json') as f:
            config = json.load(f)
        
        #Parsing
        if len(config['jsonFile']) > 0:
            print('Load Config Success!')
            return config
        else:
            print('No Input')
            return False

    def loadJsonFile(self, files):
        dataDict = {}
        for fn in files:
            with open('./inputFile/' + fn + '.json', encoding="utf-8") as f:
                print('Load Json Success: '+ fn)
                data = json.load(f)
                dataDict[fn] = data
        return dataDict
    
    def loadExcelFile(self, files):
        dataAry = {}
        for fn in files:
            dataAry = pd.read_excel('./inputFile/' + fn + '.xlsx',index_col=0).to_dict()
            print('Load Excel Success: '+ fn)
        return dataAry

    def parsingJsonFile(self, files, inputType):
        print('Start Parsing File!')
        dataDict = {}
        returnDict = {}
        keyAry = []
        sortedKeyAry = []


        if(inputType == 'json'):
            dataDict = self.loadJsonFile(files) 
        elif inputType == 'excel':
            dataDict = self.loadExcelFile(files) 

        for data in dataDict:
            returnDict[data] = {}
            for key in dataDict[data]:
                if key not in keyAry:
                    keyAry.append(key)

        sortedKeyAry = sorted(keyAry, key=str.lower)#key=str.lower 大小寫一起做排序
        for data in returnDict:
            for key in sortedKeyAry:
                if key not in dataDict[data]:
                    returnDict[data][key] = ""
                else:
                    value = dataDict[data][key]
                    bl = pd.isna(value)
                    if bl:
                        returnDict[data][key] = ""
                    else:
                        returnDict[data][key] = value
        
        print('End Parsing File!')
        return returnDict,sortedKeyAry

    def exportToExcel(self,data,keyAry):
        print('Start Export File!')
        
        dataframe = pd.DataFrame(data,index=keyAry)
        file_name = './outputFile/excel.xlsx'
        
        dataframe.to_excel(file_name)
        print('End Export File!')
    
    def exportToJson(self,data):
        print('Start Export File!')
        for key in data:
            with open('./outputFile/'+ key + '.json', 'w', encoding="utf8") as f:
                json.dump(data[key], f, ensure_ascii=False)
        print('End Export File!')


jsonCon = jsonConvertor()

