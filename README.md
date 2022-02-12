# csvLogToExcel
The main task of this prject is to convert some logs in one excel files.

The log files are under the folder 
/MarketReports
    /dd_mm_yyyy
        /Logs
        /MatchedBets
        /ProfitAndLoss

The proupuse is to merge the log with the same name and generate some rows in the excel SPORT REPORT.xlsx

Is so important to mantain the structure of the excel, the columns where place the data and the formattation.

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
            MacthedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv

That 3 file logs the same match.

after that ONLY IF "MacthedBetsReport" have some row, like this:

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
            MacthedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv
        /ProfitAndLoss
            ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv

Form "Log_10_02_2022_Sasnovich v Cristian - Match Odds.csv"

We create a row live that:
![main info](.\images\1.png?raw=true "main info")

Were we have 
note that always have that data:

    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for Aliaksandra Sasnovich: info = marketName: Sasnovich v Cristian - Match Odds; runnerA: Aliaksandra Sasnovich; runnerB: Jaqueline Cristian 

And that row

    2/10/2022 14:21:58: [G_Auto 2] :  Store Value (Shared) for market: aId = 7283310
    
    2/10/2022 14:21:58: [G_Auto 2] :  Store Value (Shared) for market: bId = 9627961
    
    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for market: inplayInfo = aBsp: 1.6; bBsp: 2.66; volume: 105741.4 
    
    2/10/2022 14:21:58: [G_Auto 2] :  Store Text Value (Shared) for market: inplayInfo = aBsp: 1.6; bBsp: 2.66; volume: 105741.4 


![main info2](.\images\2.png?raw=true "main info2")

It important to maintain always the column identation.

# 2- Add the profit and loss of the market

Now form the files "ProfitLossReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv"

we Take the "Profit If Win" and "Profit If Lose" column

```csv
Time, Selection, Profit If Win, Profit If Lose
15:29:45,Aliaksandra Sasnovich,-2.50,0.00
15:29:45,Jaqueline Cristian,45.75,0.00
```

and place that in colums AG and AH, respectively in the same row of market name for runner A and below for runner B

![main info3](.\images\3.png?raw=true "main info3")


# 2- Add the mathched bets row

Always star for the row where the market name are, form column Z to column AE we place the matched bets, in CRESCENT order of time

We take that info form "MacthedBetsReport_10_02_2022_Sasnovich v Cristian - Match Odds.csv" files

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
- for the first in column AA we put "OPEN"
- in column AB we put the A if the Selection in csv file is the same of runnerA, B otherwise.
- We place L if Back/Lay in csv is == to Lay, B if == Back
- In colum AD we put the "Odds"
- In column AE we put the "Stake"
- after that row we put an empy row were we wirte in column "AA" -> "FINAL"

note, we need in colum Y the id of the but as the sequencil number of the bets

![matched bets](.\images\4.png?raw=true "matched bets")

NB: if we have, like in that example a duplicate row, with the SAME time, we just one row in excel and we place as a stake the sum of the two
STAKE = 6.94+6.94

```csv
Time, Back/Lay, Selection, Odds, Stake, Reference, SP Bet
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113995,
10/02/2022 13:37:22,Lay,Aliaksandra Sasnovich,1.36,6.94,258207113062,
```
![matched bets2](.\images\5.png?raw=true "matched bets2")


