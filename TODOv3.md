# TODO

add time for aFirstSetPrice and aSecondSetPrice, check the time of the row and transform that in UTC time stamp millisecond 14 digits (as explained in excelToJson todo)

aFirstSetPrice must be placed in in column B row 2



![new data1](./documentation//images/4/1.png?raw=true "new data1")

if is present aSecondSetPrice place that in column B row 3

![new data2](./documentation//images/4/4.png?raw=true "new data2")

# new params

i need to copy more params from the log

all the data will be called

For runner A:
aParams1
aParams2
...
aParams10

For runner B:
bParams1
bParams2
...
bParams10

not all params will be present in the logs, that will be the output in logs:


```csv
2/28/2022 18:40:28: [G_Auto 1] :  Store Text Value (Shared) for Nicolas Jarry: aParams1 = 3.4
2/28/2022 18:40:28: [G_Auto 1] :  Store Text Value (Shared) for Matheus Pucinelli De Al: bParams1 = 1.41

...

2/28/2022 19:45:22: [G_Auto 1] :  Store Text Value (Shared) for Nicolas Jarry: aParams2 = 5.4
2/28/2022 19:42:22: [G_Auto 1] :  Store Text Value (Shared) for Matheus Pucinelli De Al: bParams2 = 1.2
```

for each params present copy that in ow 1 (the same of marketName) and  columns AL to AU for runnerA (in that case Nicolas Jarry) and in columns AV to BE for runnerB

as example:
![new data2](./documentation//images/4/2.png?raw=true "new data2")


to test the features use mainStrategy2.py and the only market that have the params in the logs are "Log_19_03_2022_Halep v Swiatek - Match Odds.csv" in folder "Strategy2MarketReports", the result file are /Strategy2ExportReport/REPORT_20_03_2022_18_27.xlsx

that are the example
![new data3](./documentation//images/4/3.png?raw=true "new data3")





