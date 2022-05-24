# VBA Stock Analysis
---
## Purpose 
This VBA script was created to analyze and output stock data from 2017 and 2018 and display the ***Total Daily Volume*** for the year and the ***Return Percentage*** for each stock. By looking at the increase or decrease in a stocks Return Percentage, someone could determine what a good stock to invest in for the following year could be. 

There is also some refactoring of code to increase the speed and effeciency of the output

![STOCKS](https://user-images.githubusercontent.com/60283799/170139665-4fb38c18-3617-4089-9966-390d6e9dca58.jpeg)

## Results 

---

### Analysis for 2017 stocks 

![2017 Message Box](https://user-images.githubusercontent.com/60283799/170143710-cca36925-8cee-4c0a-88e0-7f6200ac5837.PNG)

In 2017, the market for these specific stocks was quite good. With an average ***Total Daily Volume*** of $263,886,592 and an average ***Return Percentage*** of 67.3%, it was a positive overall year for any investors of this group of stocks. Many stocks averaged a more than 100% return rate which is very remarkable. This could be an indicator as to what happened with the 2018 stocks. 

---

### Analysis for 2017 stocks  

![2018 Message Box](https://user-images.githubusercontent.com/60283799/170143852-911d78ca-6f80-4fd5-a191-842f24d3d74e.PNG)

2018 was a much harsher year for these specific stocks. The average ***Total Daily Volume*** remained about the same at $275,503,183, however the average ***Return Percentage*** was significantly lower at around -8.5%. It makes me wonder whether it was a Market Correction, and the stocks simply had a good year in 2017 and reverted back to "normal" in 2018, but with such a large difference it could be many more factors such as the status of the economy or the fact that it was a mid-term election year. 

---

### Code Comparison and Run Time 

Originally the code was written a nested for loop using two loops to run through and anylize the data. 

The first one to loop shown below was to loop through each ticker and reset the volume back to 0 and move on to the next i or next ticker. 
```
For I = 0 To 11
       ticker = tickers(I)
       totalVolume = 0
```

After looping through the first loop, the second nested loop would do the grunt work, looping through the tickers, storing the Daily Volume and Return, and then outputting it to the cell one ticker at a time. This works, however it slows the code down by having to do the process one ticker at a time. The second loop is shown below. 


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

In Total, this took around 20 seconds for the analysis to complete. <br />
Below is an image of both runtimes before refactoring. 

<br />

![SLOW 2017 message](https://user-images.githubusercontent.com/60283799/170148808-743bd8a4-afd4-42f4-86e1-89c22bbc1de0.PNG)
![SLOW 2018 message](https://user-images.githubusercontent.com/60283799/170148812-cf33d7db-1b23-4ddc-a2be-b2fe16d8676f.PNG)



## Summary 































