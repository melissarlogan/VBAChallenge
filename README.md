# VBA Challenge
## Overview of Project
### Background
Steve’s family is looking to invest in the stock market within green energy stocks specifically. The family; however, is only looking at one stock “DQ”. The below project expanded that research to analyze not only the "DQ" stock but several others to gage their performance over 2017 and 2018 to allow Steve’s family to make the best investing decision.
### Purpose
The purpose of this project was to analysis multiple stocks at the push of a button to compare their year over year trends. This project also took that original analyzing code  and refactored it to improve the code to make it faster. Refactoring makes the code run more efficiently using less steps or less memory and makes the code easier for anther developer to review and make changes to. By refactoring the code, it makes it easier on Steve’s computer to run the analysis and allows for more stocks to be added to the code in the future to analyze.

## Results 
### 2017
In the below image the 12 stocks are shown comparing their Total Daily Volume as well as  the change in price from the first trade in 2017 compared to the last trade in 2017.  The last and first trades share prices are pulled as follows within VBA.

Starting Price
```
If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
```

Ending Price
```
If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
```
This formula is saying within the data provided, if the row before it is for the same ticker then there is still an earlier stock price and same for the ending price if there is a row with the same ticker afterwards. This code works given the list  we have is already in order by the stock tickers by days of the year it was traded.

The data shows that stocks “FSLR” and  “SPWR” were the most traded stocks of this sample in 2017. The data also shows that in 2017 the stocks with the highest positive change in their values were “DQ” and “SEDG” both with returns over 180%. Vise vera we can see that “DQ” and “HASI” had the lowest volume of trades and “RUN” and “TERP” had the worst performance.  “DQ” is the stock that Steve’s parents were looking into, so while it did have one of the highest returns in 2017, it does not have the high trading volume that his parents were looking for. Overall, it should be noted that all but 1 stock in this year had a positive return, this can be a result of macro trends in the economy implying that the green energy sector was up in 2017.

![2017 Results](https://user-images.githubusercontent.com/85718354/124306803-e879a280-db34-11eb-8600-2991b96032b7.JPG) 

This code was run in 1.21 seconds before the refactoring and in 0.15 seconds after refactoring.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/85718354/124848744-74dbf900-df6b-11eb-9aec-1c729664d600.JPG)



![VBA_Challenge_2017](https://user-images.githubusercontent.com/85718354/124310653-815eec80-db3a-11eb-95f5-b3c32f9caec5.JPG)

### 2018
In the next image this is showing the same 12 stocks as the 2017 analysis but for a more recent year (2018). As shown in the image, 2018 was overall a worse year for these stocks given majority are showing red in the return column, meaning the price the company first traded within the year is lower than what it sold for in its last trade of the year. 

Return each year is calculated using the below code taking the difference between the beginning and ending prices.
```
Cells(4 + i, 3).Value = ((tickerEndingPrices(i) / tickerStartingPrices(i)) - 1)
```

Top performers in terms of volume were “ENPH” and “SPWR” and performers in return were “ENPH” and “RUN”. Worst performers in volume were “AY” and “HASI” and for return were “DQ” and “JKS”. This shows a large contrast within “DQ” as this was the stock his parents were interested in, rising 199.4% in 2017 and then falling 62.6% in 2018. At a macro level all but 2 stocks had a negative return this year indicating that the green energy sector was down.

Red and Green within the data set are assigned based on amounts greater than or less than zero using the following code. This is telling VBA if the return on the prices is greater than 0, meaning the ending price is higher than the starting price, to colour the cell green, and red if the ending price is lower than the starting. 
```
If Cells(i, 3) > 0 Then
Cells(i, 3).Interior.Color = vbGreen
Else: Cells(i, 3).Interior.Color = vbRed
End If
```

![2018 Results](https://user-images.githubusercontent.com/85718354/124307473-dfd59c00-db35-11eb-82bf-d9f2ce72d4f2.JPG)

This code was run in 1.20 seconds before refactoring and in 0.19 after refactoring

![VBA_Challenge_2018](https://user-images.githubusercontent.com/85718354/124848836-a81e8800-df6b-11eb-865d-761b7244841f.JPG)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/85718354/124310704-950a5300-db3a-11eb-88fe-be248df1c28b.JPG)


## Summary
### Refactoring Code Generally
Code refactoring is taking functioning code and modifying it. It is done with the purpose of preserving the functionality of the code while making the code run faster, run more efficient, or be written clearer in the event the code needs to be changed or used elsewhere.
#### Advantages of Code Refactoring
1.	Refactoring can improve code readability by revising the code and determining how to make it easier to read by the end user, this could include comments in the code or adding/removing unnecessary variables to make the code clearer.	
2.	Refactoring can reduce code complexity by removing unnecessary functions or achieving the same output but using fewer commands.
3.	Refactoring can increase the speed in which the code is executed  by requiring less memory or capacity from the user’s computer through removing complexities and improving the execution.
#### Disadvantages of Code Refactoring
1.	Refactoring takes time  which can be seen as unrequired given the code is already working.
2.	Refactoring can be tedious as a small change made could cause bugs later down the line that will need to be fixed.
### Original and refactored VBA Script
Within the stock analysis workbook, the original code has been refactored to run easier on the computer and be able to scale easier. This is done by introducing more variables per the below code and assigning those variables within the code to function per the below:

```
'Initialize array of all tickers so that they can be analyzed.
Dim tickers(12) As String
    
tickers(0) = "AY"
tickers(1) = "CSIQ"
tickers(2) = "DQ"
tickers(3) = "ENPH"
tickers(4) = "FSLR"
tickers(5) = "HASI"
tickers(6) = "JKS"
tickers(7) = "RUN"
tickers(8) = "SEDG"
tickers(9) = "SPWR"
tickers(10) = "TERP"
tickers(11) = "VSLR"

    
'Get the number of rows to loop over- finding the ending row.
RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'1a) Create a ticker Index (make the ticker index move as the tickers increase)
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```

#### Advantages of the Original Script (No Refactoring)
1. The original VBA script worked and functioned as was intended meaning any changes will not impact the final result which was what Steve has requested.
2. The original VBA ran in a reasonable amount of time to analyze 12 stocks. While the refactored script is faster, it took more time to refactor than it did to originally build.
#### Disadvantages of the Original Script (With Refactoring)
1. The original script while functioning for the 12 stocks, would be harder to add in 1000’s of stocks. With refactoring this code is now more scalable if Steve’s parents were looking to analyze more stocks. 
2.  With refactoring more indicators (comments) of what the code is doing is added making it easier to read by the end user. The original code was written by one person and would have been easy for the one person to make changes, however by taking the time to edit and improve this allows a different user to make changes and understand what is happening and going to happen in the code.


