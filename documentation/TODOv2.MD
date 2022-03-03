## add to additional row and shift th present -> row K RUNNERS SET A, row L RUNNERS SET B


Strategy2MarketReport
mainStrategy2.py

![add new rows](./images/3/1.png?raw=true "add new row")

- [ ] i added some the logs information, only in the new one logs but i need that the code works already in the older, so if it's not present the present the row in Logs, leave empty in the excel

in that example "Log_01_03_2022_Giannessi v Rosol - Match Odds.csv" we have the new line 

![new data1](./images/3/2.png?raw=true "new data1")

![new data2](./images/3/3.png?raw=true "new data2")



```csv
2/28/2022 18:40:28: [G_Auto 1] :  Store Value (Shared) for market: aFirstSetPrice = 3.4
2/28/2022 18:40:28: [G_Auto 1] :  Store Value (Shared) for market: bFirstSetPrice = 1.41
```

and 


```csv
2/28/2022 19:15:17: [G_Auto 1] :  Store Value (Shared) for market: aSecondSetPrice = 1.81
2/28/2022 19:15:17: [G_Auto 1] :  Store Value (Shared) for market: bSecondSetPrice = 2.24
```

must be copied in the ROW K for aFirstSetPrice and aSecondSetPrice
and in row L for bFirstSetPrice and bSecondSetPrice, and must start for first row and go down, like that

![new data3](./images/3/4.png?raw=true "new data3")

like i said, if that rows are not present in the log leave the columns empty.


## fix for faster bet
As i told there are a problem when the bets are so fast, i already tried to fix my logger but i cant do best.

So we need to fix that via code

I use that file as Example "REPORT_01_03_2022_18_26" in documentation folder.

Talking about the market "Log_25_02_2022_Borges v Ritschard - Match Odds.csv" 

![faster](./images/3/5.png?raw=true "faster")

What i need is that:
![faster2](./images/3/6.png?raw=true "faster2")

So when you reach a "fast bets" the have same time or the difference between are less than 5 second you have to make a copy of the first point for all the bets that have same time (or in 5 second range)


### Another example 1

"Log_25_02_2022_Borges v Ritschard - Match Odds.csv"

From this
![faster3](./images/3/7.png?raw=true "faster3")


to this
![faster4](./images/3/8.png?raw=true "faster4")


### Another example 2

"Log_25_02_2022_Hurkacz v Rublev - Match Odds.csv"

From this
![faster5](./images/3/9.png?raw=true "faster5")


to this
![faster6](./images/3/10.png?raw=true "faster6")

### Another example 3

"Log_25_02_2022_Karolina Schmiedlova v Q Wang - Match Odds.csv"

From this
![faster7](./images/3/11.png?raw=true "faster7")


to this
![faster8](./images/3/12.png?raw=true "faster8")

