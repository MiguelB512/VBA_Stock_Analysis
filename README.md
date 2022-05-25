# VBA Stock Analysis
---
## Purpose 
This VBA script was created to analyze and output stock data from 2017 and 2018 and display the ***Total Daily Volume*** for the year and the ***Return Percentage*** for each stock. By looking at the increase or decrease in a stocks Return Percentage, someone could determine what a good stock to invest in for the following year could be. 

There is also some refactoring of code to increase the speed and effeciency of the output.

![STOCKS](https://user-images.githubusercontent.com/60283799/170139665-4fb38c18-3617-4089-9966-390d6e9dca58.jpeg)

## Results 

### Analysis for 2017 stocks 

![2017 Message Box](https://user-images.githubusercontent.com/60283799/170143710-cca36925-8cee-4c0a-88e0-7f6200ac5837.PNG)

In 2017, the market for these specific stocks was quite good. With an average ***Total Daily Volume*** of $263,886,592 and an average ***Return Percentage*** of 67.3%, it was a positive overall year for any investors of this group of stocks. Many stocks averaged a more than 100% return rate which is very remarkable. This could be an indicator as to what happened with the 2018 stocks. 

---

### Analysis for 2017 stocks  

![2018 Message Box](https://user-images.githubusercontent.com/60283799/170143852-911d78ca-6f80-4fd5-a191-842f24d3d74e.PNG)

2018 was a much harsher year for these specific stocks. The average ***Total Daily Volume*** remained about the same at $275,503,183, however the average ***Return Percentage*** was significantly lower at around -8.5%. It makes me wonder whether it was a Market Correction, and the stocks simply had a good year in 2017 and reverted back to "normal" in 2018, but with such a large difference it could be many more factors such as the status of the economy or the fact that it was a mid-term election year. 

---

## Code Comparison and Run Time 

Originally the code was written a nested for loop using two loops to run through and analyze the data. 

The first one to loop shown below was to loop through each ticker and reset the volume back to 0 and move on to the next i or next ticker. 
```
For I = 0 To 11
       ticker = tickers(I)
       totalVolume = 0
```

After looping through the first loop, the second nested loop would do the grunt work, looping through the tickers, storing the Daily Volume and Return, and then outputting it to the cell one ticker at a time. This works, however it slows the code down by having to do the process one ticker at a time. The second loop is shown below:


```
       For J = 2 To RowCount
       Sheets(yearValue).Activate
       
           '5a) Get total volume for current ticker
           
           If Cells(J, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(J, 8).Value

           End If
           
           '5b) get starting price for current ticker
           
           If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then

               startingPrice = Cells(J, 6).Value

           End If

           '5c) get ending price for current ticker
           
           If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then

               endingPrice = Cells(J, 6).Value

           End If
       
       Next J
       
       '6) Output data for current ticker
       
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + I, 1).Value = ticker
       Cells(4 + I, 2).Value = totalVolume
       Cells(4 + I, 3).Value = endingPrice / startingPrice - 1

   Next I
```
<br />

In total, this took around 20 seconds for the analysis to complete. <br />
Below is an image of both runtimes before refactoring. 

<br />

![SLOW 2017 message](https://user-images.githubusercontent.com/60283799/170148808-743bd8a4-afd4-42f4-86e1-89c22bbc1de0.PNG)
![SLOW 2018 message](https://user-images.githubusercontent.com/60283799/170148812-cf33d7db-1b23-4ddc-a2be-b2fe16d8676f.PNG)

<br />


The refactored code is similar but has a few changes to improve the efficiency. Instead of having the outputs be posted one by one as the code runs, the loop instead stores the values in the tickerIndex and outputs it at the end of the loop. Although there is still two loops, these are not nested loops and also the values are not being outputed directly after the loop, which is less computing for the script. The code is shown below:

```
For I = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         
         If Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
            
        'End If
        
         End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(I, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        'End if
        
        End If
    
    'Loop to next i
    
    Next I
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For J = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + J, 1).Value = tickers(J)
        Cells(4 + J, 2).Value = tickerVolumes(J)
        Cells(4 + J, 3).Value = (tickerEndingPrices(J) / tickerStartingPrices(J)) - 1
        
    Next J
```

<br />

This brought the run time down to about half of a second which is a 3233.3% decrease in time and would be quite beneficial if the source data was larger. 
Below is a screenshot of the two run times after refactoring. 

![FAST 2017 message](https://user-images.githubusercontent.com/60283799/170163717-4818a847-4512-4ae4-8944-a6242c0e0981.PNG)
![FAST 2018 message](https://user-images.githubusercontent.com/60283799/170163714-9e187d8d-8837-41fe-9a5e-5c55fcbaf69c.PNG)


## Summary 

Code effeciency and order is very important in programming. A 3233.3% increase in speed is infinitely useful and always beneficial. Assuring that code is as smooth and concise as can be while still achieving the goal is the epitome of programming. 

When it comes to refactoring, there are both pros and cons. 

Some Pros are 
- Coding can become much quicker thanks to refactoring code that someone has previously written. 
- It can inspire you to change the code to be even quicker, more specific, or just be the missing puzzle piece that was needed
- For new coders, it's a good way to learn and experiment to fully understand how to properly or efficiently write code

Some of the cons are 
- If not detailed, it can be confusing to edit code that was written by someone else
- If not appreciated, a programmer may just copy and paste without fully understanding the code and being able to implement it on their own 

<br />

When it comes to refactoring the code for this project, i didn't encounter too many of these pros and cons due to the fact that it was code that was written by me throughout the course of the module. I did learn alot from searching the internet and reading about for loops, conditionals and much more which was a huge help. This exercise taught me to not be complacent with the finished product because it could always be better. Refactoring will definitely be something that i use in future projects, not only code that i find online but also with code that i had previously written myself. 




























