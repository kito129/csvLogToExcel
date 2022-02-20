# Csv Log To Excel
The main task of this project is to convert some logs in one excel files.

I need a python script to run every day over some csv logs.

The log files input will be in these folders 

    /MarketReports
        /dd_mm_yyyy
            /Logs
            /MatchedBets
            /ProfitAndLoss

Please set a variable in python that i can change, that refer to path of /MarketReports, so i need to run over different path is possible to do that

    path = "./MarketReports/"
    exportPath = "./ExportReport/"

The proposes is to merge the log with the same name and generate some rows in the excel SPORT REPORT_dd_mm_yyyy_hh_mm.xlsx with the date and time when the scripts has been launched.

After the scripts run, I need all the market found in the logs folder with AT LEAST one matched bet to be  in the excel

Is so important to maintain the structure of the excel, the columns where place the data and the format of the excel.
SPORT REPORT.xlsx are here as example

# 1- Check for the file 
First of all you need to look up over every folder

    /dd_mm_yyyy
        /Logs
        /MatchedBets
        /ProfitAndLoss

and match the files, as example we have

    /10_02_2022
        /Logs
            Log_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv

That 3 file logs the same match.

after that ONLY IF "MatchedBetsReport" have some row, like this:

```csv
Time, Selection, Profit If Win, Profit If Lose
15:29:45,Aliaksandra Sasnovich,-2.50,0.00
15:29:45,Jaqueline Cristian,45.75,0.00
```

so the market must be added in the excel.

If the file are like that:
```csv
Time, Selection, Profit If Win, Profit If Lose
```
We just skip that market.

NOTE: File that present this in the name: Log_11_02_2022_Market Closed (x.xxxxxxxxx) must be skipped.

# 1- Add the main info of the market

For the market above we need to add a row in the excel.

Taking always as reference

    /10_02_2022
        /Logs
            Log_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv

Form "Log_10_02_2022_Sasnovich v Cristian - Match Odds.csv"

We create a row live that:
![main info](/images/1.png?raw=true "main info")

Were we have 
note that always have these data:

    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: info = marketName: Sasnovich v Cristian - Match Odds; runnerA: Aliaksandra Sasnovich; runnerB: Jaqueline Cristian 

And these row

    2/10/2022 14:21:58: [G_Auto 2] :  Store Value (Shared) for market: aId = 7283310
    2/10/2022 14:21:58: [G_Auto 2] :  Store Value (Shared) for market: bId = 9627961
    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for market: inplayInfo = aBsp: 1.6; bBsp: 2.66; volume: 105741.4 
    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for market: inplayInfo = aBsp: 1.6; bBsp: 2.66; volume: 105741.4 


![main info2](/images/2.png?raw=true "main info2")

It important to maintain always the column indentation.

# 2 - Add the profit and loss of the market

Now from the file "ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv"

we Take the "Profit If Win" and "Profit If Lose" column
 and place that in colum AG and AH in the first row (the same of the market name) for runnerA and in the second row for runnerB
```csv
Time, Selection, Profit If Win, Profit If Lose
15:29:45,Aliaksandra Sasnovich,-2.50,0.00
15:29:45,Jaqueline Cristian,45.75,0.00
```

and place that in columns AG and AH, respectively in the same row of market name for runner A and below for runner B

![main info3](/images/3.png?raw=true "main info3")


# 3 - Add the matched bets row

Always start for the row where the market name are, form column Z to column AE we place the matched bets, in CRESCENT order of time ( you need to reverse the order of the logs)

We take that info form "MatchedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv" files

```csv
Time, Back/Lay, Selection, Odds, Stake, Reference, SP Bet
10/02/2022 14:25:43,Lay,Aliaksandra Sasnovich,1.05,19.99,258210913834,
10/02/2022 14:07:26,Back,Aliaksandra Sasnovich,1.2,13.88,258209447915,
10/02/2022 14:05:50,Back,Aliaksandra Sasnovich,1.15,16.66,258209329033,
10/02/2022 14:03:38,Lay,Aliaksandra Sasnovich,1.08,31.24,258209131668,
10/02/2022 13:57:03,Lay,Aliaksandra Sasnovich,1.09,27.77,258208704386,
10/02/2022 13:51:36,Back,Aliaksandra Sasnovich,1.17,25.64,258208304505,
10/02/2022 13:48:04,Lay,Aliaksandra Sasnovich,1.2,25,258207972263,
10/02/2022 13:42:46,Back,Aliaksandra Sasnovich,1.28,40.58,258207610316,
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113995,
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113062,
10/02/2022 13:30:43,Back,Aliaksandra Sasnovich,1.44,5.68,258206634020,
10/02/2022 13:28:48,Lay,Aliaksandra Sasnovich,1.32,31.24,258206491783,
```

- We put the time, in column Z
- for the first bets, in column AA we put "OPEN"
- in column AB we put the A if the Selection in csv file is the same of runnerA, B otherwise.
- We place L if Back/Lay in csv is == to Lay, B if == Back
- In colum AD we put the "Odds"
- In column AE we put the "Stake"
- after that row we put an empty row were we write in column "AA" -> "FINAL"


we need in colum Y the id of the but as the sequential  number of the bets

![matched bets](/images/4.png?raw=true "matched bets")

NOTE 1: if we have, like in that example a duplicate row, with the SAME time, we add only one row in excel and we put as a "stake" the sum of the two stake
example:
STAKE = 6.94+6.94

```csv
Time, Back/Lay, Selection, Odds, Stake, Reference, SP Bet
....
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113995,
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113062,
.....
```
![matched bets2](/images/5.png?raw=true "matched bets2")


NOTE 2: if we have only one matched bets add final in a below row to maintain as the last value for the market.

![final value](/images/10.png?raw=true "final value")



# 4 - Add the tennis point for each matched bets

Now thats the matched bets are in the excel we need to find the tennis point when each bets were placed.

That info are under the file Logs.
Continuing our example on "Log_10_02_2022_Sasnovich v Cristian - Match Odds.csv"

we are looking to find every block like that (row 9 to 20 in the log)

```csv
...
2/10/2022 14:28:46: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: aPoints = point: 0; games: 2; sets: 0 
2/10/2022 14:28:46: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: aServing = serving:1
2/10/2022 14:28:46: [G_Auto 2] :  Store Text Value (Shared) for Jaqueline Cristian: bPoints = point: 0; games: 1; sets: 0 
2/10/2022 14:28:46: [G_Auto 2] :  Store Text Value (Shared) for Jaqueline Cristian: bServing = serving:0
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: price = 1.32
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: totalBack = 0
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: totalLay = 31.24
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: stake = 31.24
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: totalStake = 31.24
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: pl = -0.24
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: currentOddsA = 1.32
2/10/2022 14:28:46: [G_Auto 2] :  Store Value (Shared) for market: currentOddsB = 4.1
...
```

Each block of the log like that contains the point information.

So just put that data in columns from L to W

First Two rows are referring to runnerA, and the second two to runnerB

So for the first bet we have 
runnerA sets=0 and runnerB sets=0, so we are in the first set -> place games values in colum L and M:
Under game A and B (columns V and W) we place the point:
runnerA: 0
runnerB: 0

For SER colum (X columns) we place A if aServing =1 and bServing = 0 , B if bServing = 2 and aServing = 0.

Keep in mind that other info it's only to refer to the bets, so for example price =1.32 is for the first matched bet @ 1.32 added in row 4 of the excel.

![tennis point](/images/6.png?raw=true "tennis points")

For each matched bets added we add the tennis point

## 4.1 - Repeated point logs
Is possible in the logs to have two or more repeated info, like in the example:

Block form row 45 - 56 and block 57 - 68 refer to the same matched bets, but as you can see the value TOTAL STAKE (row 53 and 65) are not changed, so isn't a new matched bet but the same. 

Always take the info from the first block. 

Always check if the new block have the same TOTAL STAKE, so it refer to the same matched bets, even if other values are changed (time or point or other values).

![tennis point2](/images/7.png?raw=true "tennis points2")

## 4.2 - Another points example
Another example for bet id 6 (row 9 of the excel):

we have these block (row 81 to 92):

```csv
2/10/2022 14:51:33: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: aPoints = point: 0; games: 0; sets: 1 
2/10/2022 14:51:33: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: aServing = serving:0
2/10/2022 14:51:33: [G_Auto 2] :  Store Text Value (Shared) for Jaqueline Cristian: bPoints = point: 0; games: 0; sets: 0 
2/10/2022 14:51:33: [G_Auto 2] :  Store Text Value (Shared) for Jaqueline Cristian: bServing = serving:2
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: price = 1.17
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: totalBack = 71.9
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: totalLay = 70.12
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: stake = 25.64
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: totalStake = 142.02
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: pl = -1.77
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: currentOddsA = 1.17
2/10/2022 14:51:33: [G_Auto 2] :  Store Value (Shared) for market: currentOddsB = 6.6
```
The point are now 
aSet: 1, bSet:0, so we are in second set:

we gonna put in columns L and M (first set) 1 and 0, and in column N and O (second set) the current game point, so: 0 and 0

The same for third and eventually fourth or fifth set

![tennis point3](/images/9.png?raw=true "tennis points3")

NOTE 1: if the games are 6 for runnerA and runnerB we are in the tie break, and the log don't put the point, just leave at 0 the values for the columns V and W

NOTE 2: if aPoint: == 99 or bPoint: == 99 we have an Advantage for the runner so we gonna put "AD" for the runner that have 99 as a value, and 40 for the runner that have 40

## 4.3 - Another points example 2

![tennis point4](/images/11.png?raw=true "tennis points4")

In that example we are in the third set, so column L and M for the first and N and O for the second are set to 1 1 and 1 1, and we are 1-1 in the third

![tennis point5](/images/12.png?raw=true "tennis points5")


# 5 - Final result for the example

As a final result for these example we would have a excel like that

![final result](/images/8.png?raw=true "final result")


In the SPORT REPORT.xlsx you will find other 4 examples that are referring to these matches:

    1) row 16 to 20
    /11_02_2022
        /Logs
            Log_11_02_2022_Ostapenko v Sasnovich - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_11_02_2022_Ostapenko v Sasnovich - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_11_02_2022_Ostapenko v Sasnovich - Match Odds.csv

    2) row 21 to 24
    /11_02_2022
        /Logs
            Log_11_02_2022_Van Assche v Lestienne - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_11_02_2022_Van Assche v Lestienne - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_11_02_2022_Van Assche v Lestienne - Match Odds.csv

    3) row 25 to 27
    /12_02_2022
        /Logs
            Log_12_02_2022_Ostapenko v Kontaveit - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_12_02_2022_Ostapenko v Kontaveit - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_12_02_2022_Ostapenko v Kontaveit - Match Odds.csv

    4) row 28 to 30
    /12_02_2022
        /Logs
            Log_12_02_2022_Tsitsipas v Lehecka - Match Odds.csv
        /MatchedBets
            MatchedBetsReport_12_02_2022_Tsitsipas v Lehecka - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_12_02_2022_Tsitsipas v Lehecka - Match Odds.csv


If you need some more specification or have some problem let free to contact me at every time!

Thank you








