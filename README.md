# Stock-Analysis
### An analysis of green energy stocks for a financial advisor

## Overview

The purpose of this analysis is to show the yearly returns of a selection of green energy stock. My client is looking to advise his parents on stock options, and they have decided to invest in green energy companies. Their favorite stock is DQ, but they chose this for sentimental reasons, so are there other stocks that would give them a better financial return? 

To determine this, I want to run a basic analysis over multiple stocks. This has the added benefit of understanding looped code, as well as refactoring code for faster results. Faster results are important because our data sample right now is small, but the user could want to plug in exponentially larger data sets, which take longer to loop through. 

## Results 

There are two sides to the results: The performance of the stocks, and how the refactoring improved the code. 

### Code Breakdown 
I'll start with the code. The initial way I ran this code was with a nested `For` loop that first iterated the array values of the tickers, and then over all of the rows of data. 

```
'Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearValue).Activate
'Loop through rows in the data.
        For j = 2 To RowCount
 'Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
'Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
'Find the ending price for the current ticker.
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1) <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
 ```
This made sense to me, because I needed to move through both the array and the rows of data. However, nested loops increase the iterations exponentially. 
In my data set, I only had 12 stocks, each with many rows of data. My fictional client was thinking about analyzing all stocks using this code, which would cause the loop iterations to skyrocket. It only took a little over a second to run this code, but for more and more loops, there is a longer waiting time. 
This code ran in 1.25 seconds for 2017 and 1.23 seconds for 2018.

![Original Code Run Time - 2017](https://github.com/caseykotowski/Stock-Analysis/blob/main/Resources/Code%20Timer%20-%202017%20Original.png)

![Original Code Run Time - 2018](https://github.com/caseykotowski/Stock-Analysis/blob/main/Resources/Code%20Timer%20-%202018%20Original.png)

### Refactoring

I want to reduce the runtime, and since extra loops cause the extra time, I need to reduce the loops. I want to remove the outer loop nest from the original code, because I can make the other variables arrays, and increase the index within the remaining loop. 

```
 'Loop through all arrays to set them at zero for the start of each iteration
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
        
    'Looping over all rows
    For i = 2 To RowCount
          
        'The ticker Volume increases by the amount in the volume cell if the tickers match
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Finding where the different tickers start
         If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        'Finding where the different tickers end
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If
         
            'Increase the ticker index to switch stocks for the next loop - key for refactoring the nested loop
          If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
         End If
    
    Next i
```
I have 3 more indexed variables, which I set to zero for all indicies. From there, I only have to loop over all of the rows of data. I check for relavent volume to add, and if they're on the first or last row of data for the stock the index is connected to. 
Then, if it's the last cell for the index, I increase the ticker index by 1 to switch to the next stock. This is how we loop through all of the stocks without creating another layer to the loops. 
This increased the speed of the code considerably, while delivering the same computational results. 
For 2017, the code ran in 0.27 seconds, and 0.27 seconds for 2018.

![Refactored Code Run Time - 2017](https://github.com/caseykotowski/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![Refactored Code Run Time - 2018](https://github.com/caseykotowski/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

### Stock Results

You can see the stock performance in the screenshots of the refactored code results. 2018 was not a great year for returns on green energy stocks. DQ (the preferred stock of the parents of the user of my code) took a nearly 63% dive from 2017 to 2018, but it was up nearly 200% in 2017. I would hold off on investing in something that volitile. There are two stocks that grew in both 2017 and 2018: **ENPH & SEDG**

Those are the only stocks I would reccommend investing in based off of the numeric data over 2 years. I would be interested to see the trends for these stocks over more years, and also which types of green energies those companies invest in/create. 

## Summary

At its core, this work was about the pros and cons of refactoring code. 

### Pros

When refactoring code works well, there is an increase in efficiency. When putting large amounts of data through code, efficency can be very important, especially when your code is a deliverable for a client. You want to provide a client with the best product, and if you have the timme to refactor, it could be worth it. 

Refactoring could also make code easier to read, and easier to share because it is more simple. If you're working on a team, you want your code to make sense to more people than just you. 

### Cons

In my opinion, refactoring's biggest con is the amount of time it takes. For just this small example, it took more than an additional hour to figure out how to go about removing the extra `For` loop. I knew that's what I wanted to do, but it took time to form the new arrays and loops correctly. 

This can take even more time, because it opens you up to new bugs and typos. If you use the wrong variable, it can take several minutes to find your typo/inconsistency. If you format your new code wrong, it can take just as long to debug as to write the less efficient code to begin with. 

If you have a hard deadline looming, refactoring might not be worth it. 

### Was it worth it?

In this case, I believe it was worth the time to refactor the code. Aside from the learning opportunity, efficiency was important to my client. He wants the potential to analyze all stocks, so effiecency is the biggest feature my code could offer. 

That's not to say it wasn't without challenges. By introducing more arrays, the complex variables opened me up to more typos and bugs. I had difficulties figuring out whether to use the *i* for looping or the tickerIndex variable. I went through much trial and error to get all of the details right, and the final issue that I had to pick through to find was in my conditional statements. 
For the line `If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then` , it took a long time for me to notice that I typed `Cells(i + 1)`, instead of `Cells(i + 1, 1)` . I missed the column call, and it was hard for me to see. Refactoring opens you up to little errors like those when you edit your working code. 

In the end, I believe the reduction of execution time was worth the additional effort of refactoring. 
