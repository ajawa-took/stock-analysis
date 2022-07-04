
# VBA Stock Analysis



## Overview of Project

 Visual Basic marcos analyze the trading data for a few stocks for two years, 2017 and 2018.

### Purpose 

 To help Steve help his parents pick a good investment among these stocks.

## Results

### Stock Data Produced

![Stocks data 2017.](/resources/Stocks2017.png)
![Stocks data 2018.](/resources/Stocks2018.png)

#### Recommendations

 ENPH is by far the biggest winner here: its price more than quadrupled over the two years. DQ, parents' initial choice, is right around median over the two years: 2017 was great for it, but 2018 was terrible. As a whole, these stocks looks extremely volatile. I would not advise anyone to gamble their life savings on any of them, nor even on the whole basket of them.

#### Theories

##### High Volume?

 Steve's parents' plan to invest in stocks that trade a lot isn't great: SPWR, the most traded stock of 2017 and the second most traded in 2018, is also the second worst performer, losing over 30% over the two years.

##### Volatile!

 Clearly, these stocks as a whole had a great year in 2017 and a terrible year in 2018: *when* one invested mattered much more than which stock one picked. While the economy as a whole was growing a decent steady pace, these stocks swung wildly: a third of the annual price changes are by more than a factor of two! 

### Code Efficiency

 The refactored code runs about 4 times faster, and much more consistently. The refactored macro takes 0.05078125 seconds, with multiple runs of either year producing identical runtimes. Honestly, I am suspicious of this: shouldn't this depend on what else my not-so-fast computer is doing at the same time? My original code takes about 0.20 seconds, with run-to-run variation of about 0.01sec.


## Summary

### In General

 Refactoring code improves efficiency but takes time. Sometimes it's worth it, sometimes not. Sometimes, improving the efficiency also contributes to readability; but other times efficient code is more opaque.

### This Time

 For this particular task, the time investment was minimal, since most of the refactored code was provided. However, I don't think the efficiency gain is worth the readability loss. It's easier to understand the code when each piece does one thing; I actually broke up the task into several small subroutines.

### My Original Code Does More

 The efficiency gain from refactoring itself is actually even smaller than observed, because my original code (in `Module 1` of my excel file) does more. The refactored code that we are led to produce (in `Module 2` of my excel file) relies on the sheets of input data being sorted, first by tick and then by date; it would produce garbage from scrambled data. I did not like making that assumption when I was working through the modules. My key subroutine `VolumeReturn` works just as well on disordered data; though I did not keep that up in the outer layers. The whole macro will run fine if dates are out of order. But if rows are not sorted by ticks, it will take much longer (likely 100 times longer), and will produce many duplicate (though correct) results. I think fixing that and generally cleaning up my orignial would have been more useful than filling in the few lines left out of the refactored code provided. My original code also doesn't hard-code the tickers, nor even the number of different stocks, so it is more portable. On the other hand, maybe it's more important to efficiently accomplish the task at hand than to plan for hypothetical future tasks.


















