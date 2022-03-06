from openpyxl import Workbook
from datetime import datetime
from itertools import zip_longest
import os,csv,sys
def getField(field,content,index=0):
    try:
        try:
            line = content.split(f'{field}:',index+1)
            value = line[index+1].split(';')[0].split(' - Match Odds')[0].split('\n')[0].strip()
            if 'marketName' == field:
                dateString = line[index].split('\n')[-1].split(' ')[0]
                return value,dateString
            else:
                return value
        except:
            return content.split(f'{field} = ',index+1)[index+1].split('\n')[0].strip()
    except:
        return None
path = "./Strategy2MarketReports/"
exportPath = "./Strategy2ExportReport/"
if not os.path.exists(exportPath):
    try:
        os.mkdir(exportPath)
    except:
        print("Export Path not exists and unable to Create")
        sys.exit(0)
book = Workbook()
sheet = book.active
sheet.title = 'report'
sheet.append(["DATE","TIME","MARKET","ACCOUNT","EXECUTOR","STRATEGY","EVENT","RUNNERS","","",'RUNNER SET','RUNNER SET',"NOTE","S1 A","S1 B","S2 A","S2 B","S3 A","S3 B","S4 A","S4 B","S5 A","S5 B","GAME A","GAME B","SER","ID","TIME","TYPE","RUN","BET","ODDS","STAKE","TOTAL RISK","COMM","NET PL"])
sheet.append(["","","","","","","","A","B","WINNER",'A','B'])
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
            market['runnerSets'] = []
            market['runnerSets'].append([getField('aFirstSetPrice',content),getField('bFirstSetPrice',content)])
            market['runnerSets'].append([getField('aSecondSetPrice',content),getField('bSecondSetPrice',content)])
            pointings = {}
            i = 0
            for points in range(totalPoints):
                totalStake = getField('totalStake',content,points)
                if totalStake == getField('totalStake',content,points+1):
                    continue
                gameContent = content.split(f'(Shared) for {market["runnerA"]}: aPoints',points+1)[-1]
                aPoint = getField('aPoints = point',content,points)
                pointTime = gameContent.split(': [')[0].split('\n')[-1].split(' ')[-1]
                pointings[pointTime] = i
                aServing = getField('aServing = serving',content,points)
                aGame = getField(f'games',gameContent)
                aSet = getField(f'sets',gameContent)
                bPoint = getField('bPoints = point',content,points)
                bGame = getField(f'games',gameContent,1)
                bSet = getField(f'sets',gameContent,1)
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
                i+=1
            mList = sorted([m for m in csv.reader(mContent.splitlines()[1:])],key=lambda x: x[0].split(' ')[1])
            market['stacks'] = {}
            i=0
            stacks = []
            for row in mList:
                date = row[0].split(' ')[-1]
                flag = row[1][0]
                runner = ['A','B'][row[2].strip().lower()==market['runnerB'].lower()]
                odds = row[3]
                stack = row[4]
                if market['stacks'].get(date+"__"+odds):
                    market['stacks'][date+"__"+odds][-1]+=float(stack)
                    continue
                i+=1
                entry = None
                stacks.append(date)
                if i==1:
                    entry = 'OPEN'
                market['stacks'][date+"__"+odds] = [i,date,entry,runner,flag,odds,float(stack)]
            inserted = []
            i = 0
            for stack in market['stacks'].values():
                if stacks.count(stack[1])>1 and stack[1] not in inserted:
                    [market['sets'].insert(i+x+1,market['sets'][i]) for x in range(stacks.count(stack[1])-1)]
                    inserted.append(stack[1])
                i+=1
            final = False
            for data in zip_longest([marketDate],[marketName,market['volume']],[market['runnerA'],market['aBsp'],market['aId']],[market['runnerB'],market['bBsp'],market['bId']],market['sets'],market['stacks'].values(),market['profit'],market['runnerSets']):
                setList = []
                prList = data[6] or []
                if data[4]:
                    [setList.extend(set_) for set_ in data[4]]
                setList = setList or [None]*13
                stacks = data[5]
                if not stacks and not data[4] and not prList and not final:
                    stacks = ([None]*2)+['FINAL']+[None]*4
                    final = True
                if not stacks:
                    stacks = [None]*7
                row = [data[0],None,data[1],None,None,None,None,data[2],data[3],None]+(data[-1] or [None,None])+[None]+setList+stacks+[None]+prList
                sheet.append(row)
            if not final:
                sheet.append(([None]*26)+['FINAL'])
dateOfRun = datetime.now().strftime('%d_%m_%Y_%H_%M')
book.save(os.path.join(os.path.abspath(exportPath),f'REPORT_{dateOfRun}.xlsx'))    