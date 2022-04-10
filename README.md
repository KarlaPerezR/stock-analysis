# An Analysis of Stock Analysis

## Overview of Project
#### Analyze green stocks of the year 2017 and 2018 to determine if the stocks are worth or not investing. This Analysis was through VBA with the help of Loops but the objective of this exercise is to manage the refactor of the code to improve efficiency through arrays.

## Results

#### The better option for stock

<img align="right" src="https://github.com/KarlaPerezR/stock-analysis/blob/main/Resources/2017.png">
After the analysis of the Daily Volume and Yearly Return of the Green Stocks, the conclusion and advice is to invest in the Ticker ENPH.
<br/><br/>ENPH was the unique Ticket that had great returns for both years, 129.5% for 2017 and 81.9% for 2018, also had a good number of Daily Volume for both years, indicating that is a Stock that many people traded every day.
<br/><br/>The RUN ticket also had great returns for 2018 but had a low return in the 2017 with 5.5% so it has a volatile behavior. 
<br/><br/>And the DQ ticket, the stock that the parents of the Analyst wanted, had a great year in 2017 with a return of 199.4% but a lost of 62.6% for 2018 besides had a low Daily Volumen compare to the other stocks.
<img align="right" src="https://github.com/KarlaPerezR/stock-analysis/blob/main/Resources/2018.png">
<br clear="right"/>

#### Refactor the code

Refactor the code helps to reduce the time of processing the information, from 51,885 to 0.3125 seconds.
<br/> The creation of the arrays to save the Value of each Ticket (Total Volume, the Starting Prices and Ending Prices) allows to free memory and have a faster result.

In the first code, the values were save in Variables and then printed:

```
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
```
But with the creation of the arrays in the refactor code, the values are save in its own space in the memory:

```
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
```

## Summary: In a summary statement, address the following questions.
### What are the advantages or disadvantages of refactoring code?
### How do these pros and cons apply to refactoring the original VBA script?
