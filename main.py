from openpyxl import Workbook
from datetime import datetime, timedelta
from itertools import zip_longest
import os, csv, sys

PATH = "./MarketReports/"
EXPPATH = "./ExportReport/"

# functions 
def getField(field, content, index=0):
    try:
        try:
            line = content.split(f'{field}:', index + 1)
            value = line[index + 1].split(';')[0].split(' - Match Odds')[0].split('\n')[0].strip()
            if 'marketName' == field:
                dateString = line[index].split('\n')[-1].split(' ')[0]
                return value, dateString
            else:
                return value
        except:
            return content.split(f'{field} = ', index + 1)[index + 1].split('\n')[0].strip()
    except:
        return None

def removeOneHour(time):
    temp = datetime.strptime(time, "%H:%M:%S")
    temp -= timedelta(hours=1)
    modified_time_str = temp.strftime("%H:%M:%S")
    return modified_time_str

# main 
if not os.path.exists(EXPPATH):
    try:
        os.mkdir(EXPPATH)
    except:
        print("Export Path not exists and unable to Create")
        sys.exit(0)


book = Workbook()
sheet = book.active
sheet.title = 'report'
# columns
sheet.append(["DATE", "TIME", "MARKET", "ACCOUNT", "EXECUTOR", "STRATEGY", "EVENT", "RUNNERS", "", "", 'RUNNER SET',
              'RUNNER SET', "NOTE", "S1 A", "S1 B", "S2 A", "S2 B", "S3 A", "S3 B", "S4 A", "S4 B", "S5 A", "S5 B",
              "GAME A", "GAME B", "SER", "ID", "TIME", "TYPE", "RUN", "BET", "ODDS", "STAKE", "TOTAL RISK", "COMM",
              "NET PL", None, "aParams"] + [None] * 9 + ["bParams"])
# selections params
sheet.append([None] * 7 + ["A", "B", "WINNER", 'A', 'B'] + [None] * 25 + \
             [f'aParams{i+1}' for i in range(10)] + [f'bParams{i+1}' for i in range(10)])


for date in os.listdir(PATH):
    # files name
    datePath = os.path.abspath(os.path.join(PATH, date))
    logsPath = os.path.join(datePath, 'Logs')
    matchedBetsPath = os.path.join(datePath, 'MatchedBets')
    profitAndLossPath = os.path.join(datePath, 'ProfitAndLoss')
    logsFiles = []
    matchedBetsFiles = []
    profitAndLossFiles = []
    # iterate over directories looking for files
    for logFile in os.listdir(logsPath):
        if not logFile.startswith('Log'):
            continue
        file = logFile.split('_', 1)[1]
        if 'Market Closed' in file:
            continue
        logsFiles.append(os.path.join(logsPath, f'Log_{file}'))
        matchedBetsFiles.append(os.path.join(matchedBetsPath, f'MatchedBetsReport_{file}'))
        profitAndLossFiles.append(os.path.join(profitAndLossPath, f'ProfitLossReport_{file}'))

    # iterate over found market
    for logFile, matchedBetFile, profitAndLossFile in zip(logsFiles, matchedBetsFiles, profitAndLossFiles):
        with open(logFile) as logger, open(matchedBetFile) as matcher, open(profitAndLossFile) as profiter:
            mContent = matcher.read()
            # check if have at least 1 bets
            #if len(mContent.splitlines()) < 2:
            #    continue
            content = logger.read()
            print("File: ", logFile, end='')
            try:
                marketName, marketDate = getField('marketName', content)
            except:
                print(" (Skipped)")
                continue
            print()
            
            # get market info 
            market = {}
            market['runnerA'] = getField('runnerA', content)
            market['runnerB'] = getField('runnerB', content)
            market['aId'] = getField('aId', content)
            market['bId'] = getField('bId', content)
            market['aBsp'] = getField('aBsp', content)
            market['bBsp'] = getField('bBsp', content)
            market['volume'] = getField('volume', content)
            market['profit'] = []

            # result reader
            pReader = csv.reader(profiter.read().splitlines()[1:])
            for row in pReader:
                market['profit'].append(row[-2:])
            totalPoints = content.count('totalStake =')
            # set odds and time 
            market['sets'] = []
            d1 = content.split('aFirstSetPrice')[0].split('\n')[-1].split(': ')[0].split(' ')[-1]
            d2 = content.split('aSecondSetPrice')[0].split('\n')[-1].split(': ')[0].split(' ')[-1]

            # if time is not correct check here
            #if d1: d1 = removeOneHour(d1)
            #if d2 :d2 = removeOneHour(d2)

            market['setPriceDates'] = [None, d1, d2]
            market['runnerSets'] = []
            market['runnerSets'].append([getField('aFirstSetPrice', content), getField('bFirstSetPrice', content)])
            market['runnerSets'].append([getField('aSecondSetPrice', content), getField('bSecondSetPrice', content)])
            # selections params
            market['aParams'] = [[getField(f'(Shared) for {market["runnerA"]}: aParams{i}', content) for i in range(1,11)]]
            market['bParams'] = [[getField(f'(Shared) for {market["runnerB"]}: bParams{i}', content) for i in range(1,11)]]
            pointings = {}
            i = 0

            # iterate over bets point and serve
            for points in range(totalPoints):
                totalStake = getField('totalStake', content, points)
                if totalStake == getField('totalStake', content, points + 1):
                    continue
                gameContent = content.split(f'(Shared) for {market["runnerA"]}: aPoints', points + 1)[-1]
                aPoint = getField('aPoints = point', content, points)
                pointTime = gameContent.split(': [')[0].split('\n')[-1].split(' ')[-1]
                pointings[pointTime] = i
                aServing = getField('aServing = serving', content, points)
                aGame = getField(f'games', gameContent)
                aSet = getField(f'sets', gameContent)
                bPoint = getField('bPoints = point', content, points)
                bGame = getField(f'games', gameContent, 1)
                bSet = getField(f'sets', gameContent, 1)
                bServing = getField('bServing = serving', content, points)
                #serv = ['A', 'B'][0 if aServing else 1]
                serv = ['A', 'B'][0 if int(aServing) else 1]
                sets = [[None, None]] * 5
                if aSet == '1' and bSet == '1':
                    sets[0] = ['1', '1']
                    sets[1] = ['1', '1']
                    sets[2] = [aGame, bGame]
                elif aSet == '1' and bSet == '0':
                    sets[0] = ['1', '0']
                    sets[1] = [aGame, bGame]
                elif aSet == '0' and bSet == '1':
                    sets[0] = ['0', '1']
                    sets[1] = [aGame, bGame]
                else:
                    sets[0] = [aGame, bGame]
                aPoint = 'AD' if aPoint == '99' else aPoint
                bPoint = 'AD' if bPoint == '99' else bPoint
                if aGame == '6' and bGame == '6':
                    sets.append(['0', '0'])
                else:
                    sets.append([aPoint, bPoint])
                sets.append([serv])
                market['sets'].append(sets)
                i += 1
            # sort bets by time
            mList = sorted([m for m in csv.reader(mContent.splitlines()[1:])], key=lambda x: x[3])
            mList = sorted([m for m in mList], key=lambda x: x[0].split(' ')[1])
            # iterate over bets odds and stake
            market['stakes'] = {}
            i = 0
            betsInfo = []
            for row in mList:
                date = row[0].split(' ')[-1]
                flag = row[1][0]
                runner = ['A', 'B'][row[2].strip().lower() == market['runnerB'].lower()]
                odds = row[3]
                stake = row[4]
                if market['stakes'].get(date + "__" + odds):
                    market['stakes'][date + "__" + odds][-1] += float(stake)
                    continue
                i += 1
                entry = None
                betsInfo.append(date)
                if i == 1:
                    entry = 'OPEN'
                market['stakes'][date + "__" + odds] = [i, date, entry, runner, flag, odds, float(stake)]
            inserted = []
            i = 0

            # reorder bets
            for stake in market['stakes'].values():
                if betsInfo.count(stake[1]) > 1 and stake[1] not in inserted:
                    [market['sets'].insert(i + x + 1, market['sets'][i]) for x in range(betsInfo.count(stake[1]) - 1)]
                    inserted.append(stake[1])
                i += 1
            final = False

            # save data in excel
            for data in zip_longest([marketDate], [marketName, market['volume']],
                                    [market['runnerA'], market['aBsp'], market['aId']],
                                    [market['runnerB'], market['bBsp'], market['bId']], market['sets'],
                                    market['stakes'].values(), market['profit'], market['aParams'], market['bParams'] , market['setPriceDates'], market['runnerSets']):
                setPointList = []
                resultList = data[6] or []
                pointList = data[4]
                if pointList:
                    [setPointList.extend(set_) for set_ in pointList]
                setPointList = setPointList or [None] * 13
                betsInfo = data[5]
                if not betsInfo and not pointList and not resultList and not final:
                    betsInfo = ([None] * 2) + ['FINAL'] + [None] * 6
                    final = True
                if not betsInfo:
                    betsInfo = [None] * 7
                row = [data[0], data[-2] or None, data[1], None, None, None, None, data[2], data[3], None] + (
                            data[-1] or [None, None])
                row += [None] + setPointList + betsInfo + [None] + resultList + [None]
                row += (data[7] or ([None] * 10)) + (data[8] or ([None] * 10))
                sheet.append(row)
            if not final:
                sheet.append(([None] * 28) + ['FINAL'])


# save file               
dateOfRun = datetime.now().strftime('%d_%m_%Y_%H_%M')
outFile = os.path.join(os.path.abspath(EXPPATH), f'REPORT_{dateOfRun}.xlsx')
print("\nSaved to:", outFile)
book.save(outFile)

print("\n\n -- STATS -- \n")
print("\nLog market:", len(logsFiles))
print("Matched market:", len(matchedBetsFiles))
print("Traded market:", len(profitAndLossFiles))