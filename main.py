from openpyxl import Workbook
from datetime import datetime
from itertools import zip_longest
import os,csv,sys
def getField(field,content,index=0):
    try:
        try:
            line = content.split(f'{field}:',index+1)
            value = line[index+1].split(';')[0].split(' - ')[0].split('\n')[0].strip()
            if 'marketName' == field:
                dateString = line[index].split('\n')[-1].split(' ')[0]
                return value,dateString
            else:
                return value
        except:
            return content.split(f'{field} = ',index+1)[index+1].split('\n')[0].strip()
    except:
        return None
path = "./MarketReports/"
exportPath = "./ExportReport/"
if not os.path.exists(exportPath):
    try:
        os.mkdir(exportPath)
    except:
        print("Export Path not exists and unable to Create")
        sys.exit(0)
book = Workbook()
sheet = book.active
sheet.title = 'report'
sheet.append(["DATE","TIME","MARKET","ACCOUNT","EXECUTOR","STRATEGY","EVENT","RUNNERS","","","NOTE","S1 A","S1 B","S2 A","S2 B","S3 A","S3 B","S4 A","S4 B","S5 A","S5 B","GAME A","GAME B","SER","ID","TIME","TYPE","RUN","BET","ODDS","STAKE","TOTAL RISK","COMM","NET PL"])
sheet.append(["","","","","","","","A","B","WINNER"])
for date in os.listdir(path):
    datePath = os.path.abspath(os.path.join(path,date))
    logsPath = os.path.join(datePath,'Logs')
    matchedBetsPath = os.path.join(datePath,'MatchedBets')
    profitAndLossPath = os.path.join(datePath,'ProfitAndLoss')
    logsFiles = [];matchedBetsFiles = [];profitAndLossFiles = []
    for logFile in os.listdir(logsPath):
        if not logFile.startswith('Log'):
            continue
        file = logFile.split('_',1)[1]
        if 'Market Closed' in file:
            continue
        logsFiles.append(os.path.join(logsPath,f'Log_{file}'))
        matchedBetsFiles.append(os.path.join(matchedBetsPath,f'MatchedBetsReport_{file}'))
        profitAndLossFiles.append(os.path.join(profitAndLossPath,f'ProfitLossReport_{file}'))
    for logFile,matchedBetFile,profitAndLossFile in zip(logsFiles,matchedBetsFiles,profitAndLossFiles):
        with open(logFile) as logger,open(matchedBetFile) as matcher,open(profitAndLossFile) as profiter:
            mContent = matcher.read()
            if len(mContent.splitlines())<2:
                continue
            content = logger.read()
            print("File: ",logFile,end='')
            try:
                marketName,marketDate = getField('marketName',content)
            except:
                print(" (Skipped)")
                continue
            print()
            market = {}
            market['runnerA'] = getField('runnerA',content)
            market['runnerB'] = getField('runnerB',content)
            market['aId'] = getField('aId',content)
            market['bId'] = getField('bId',content)
            market['aBsp'] = getField('aBsp',content)
            market['bBsp'] = getField('bBsp',content)
            market['volume'] = getField('volume',content)
            market['profit'] = []
            pReader = csv.reader(profiter.read().splitlines()[1:])
            for row in pReader:
                market['profit'].append(row[-2:])
            totalPoints = content.count('totalStake =')
            market['sets'] = []
            for points in range(totalPoints):
                totalStake = getField('totalStake',content,points)
                if totalStake == getField('totalStake',content,points+1):
                    continue
                gameContent = content.split(f'(Shared) for {market["runnerA"]}: aPoints',points+1)[-1]
                aPoint = getField('aPoints = point',content,points)
                aServing = getField('aServing = serving',content,points)
                aGame = getField(f'games',gameContent)
                aSet = getField(f'sets',content,points+1 if points else points)
                bPoint = getField('bPoints = point',content,points)
                bGame = getField(f'games',gameContent,1)
                bSet = getField(f'sets',content,points+2 if points else points+1)
                bServing = getField('bServing = serving',content,points)
                serv = ['A','B'][0 if int(aServing) else 1]
                sets = [[None,None]]*5
                if aSet=='1' and bSet=='1':
                    sets[0]=['1','1']
                    sets[1] = ['1','1']
                    sets[2] = [aGame,bGame]
                elif aSet=='1' and bSet=='0':
                    sets[0] = ['1','0']
                    sets[1] = [aGame,bGame]
                elif aSet=='0' and bSet=='1':
                    sets[0] = ['0','1']
                    sets[1] = [aGame,bGame]
                else:
                    sets[0] = [aGame,bGame]
                aPoint = 'AD' if aPoint=='99' else aPoint
                bPoint = 'AD' if bPoint=='99' else bPoint
                if aGame=='6' and bGame=='6':
                    sets.append(['0','0'])
                else:
                    sets.append([aPoint,bPoint])
                sets.append([serv])
                market['sets'].append(sets)
            mList = sorted([m for m in csv.reader(mContent.splitlines()[1:])],key=lambda x: x[0].split(' ')[1])
            market['stacks'] = {}
            i=0
            for row in mList:
                date = row[0].split(' ')[-1]
                flag = row[1][0]
                runner = ['A','B'][row[2][0].strip()==market['runnerB']]
                odds = row[3]
                stack = row[4]
                if market['stacks'].get(date):
                    market['stacks'][date][-1]+=float(stack)
                    continue
                i+=1
                entry = None
                if i==1:
                    entry = 'OPEN'
                market['stacks'][date] = [i,date,entry,flag,runner,odds,float(stack)]
            for data in zip_longest([marketDate],[marketName,market['volume']],[market['runnerA'],market['aBsp'],market['aId']],[market['runnerB'],market['bBsp'],market['bId']],market['sets'],market['stacks'].values(),market['profit']):
                setList = []
                prList = data[6] or []
                if data[4]:
                    [setList.extend(set_) for set_ in data[4]]
                setList = setList or [None]*13
                stacks = data[5]
                if not stacks:
                    stacks = []
                row = [data[0],None,data[1],None,None,None,None,data[2],data[3],None,None]+setList+stacks+[None]+prList
                sheet.append(row)
            sheet.append(([None]*26)+['FINAL'])
dateOfRun = datetime.now().strftime('%d_%m_%Y_%H_%M')
book.save(os.path.join(os.path.abspath(exportPath),f'REPORT_{dateOfRun}.xlsx'))    