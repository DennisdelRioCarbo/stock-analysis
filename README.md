#  STOCK-ANALYSIS (VBA)

##  Overview of the project
    In this project we helped our friend Steve who just graduated with his finance degree to analyze green stocks.Steve wants to know the returns of some  of these green stocks to invest some of his parents money. It looks like they are particularly interested in a stock with ticker symbol "DQ".  Steve had already prepared an excel sheet with all the information for these stocks and we are going to help him automate the analysis using VBA.  
    Steve liked the workbook that was originally prepared for him using VBA in excel to automate tasks, because just with the click of a button he can analyze an entire dataset of stocks with less chance of errors. Now to make the code more efficient and be able to run the analysis for thousands of stocks, the code needs to be refactored to reduce  run times.

## Analysis And Results
    We compared 12 green stocks for Steve with the excel worksheet that he provided for us. After creating the new macro we were able to have the program analyze, return and display the ticker for each stock, calculate the total volume for that stock for each year, and calculate and display the yearly return for each stock in a format which shows wether the return was positive (green) or negative(red).
    We also created buttons to make it easier for Steve to clear the worksheet and to display all of the information for the selected year, and had VBA display the run times for the macro. This was specially useful to compare results when the code was refactored.

### To return and display the ticker for each stock we created an array with each of the stocks ticker symbol.

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

### To Calculate the total volume we also created an array and a loop that would add the daily volumes: 
       
        Dim tickerVolumes(12) As Long
        For i = 2 To RowCount
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
### We also created arrays and used conditional in the same loop as the total volume to get the starting prices and the ending prices:
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then    
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value  
        End If
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then   
         tickerEndingprices(tickerIndex) = Cells(i, 6).Value 
         End If

### We had the ticker symbol, the total volume and the return displayed on the excel sheet using loops and arrays:
       
        For i = 0 To 11
        For tickerIndex = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, 3).Value = tickerEndingprices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
         Next tickerIndex
        Next i

### We formatted the cells using conditional and cell formatting to color the inside of the cells either red or green depending on the return:

        dataRowStart = 4
        dataRowEnd = 15

        For j = dataRowStart To dataRowEnd
        
        If Cells(j, 3) > 0 Then
            Cells(j, 3).Interior.Color = vbGreen 
        Else
            Cells(j, 3).Interior.Color = vbRed
        End If
         Next j

         ### We ran a Timer to display a message on what the run time for the code was:

         startTime = Timer 
         endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

### Once we ran the timer in both the original Analysis and the Refactored Analysis we notice a decrease in run time for the refactored code for both years which indicates that the refactored code is more efficient than the original one.

**Original 2017**
![Runtime 2017 Original Stock_Analysis](https://user-images.githubusercontent.com/104289098/168509649-538acbf6-7a0b-4f0f-909c-f4a9b1c55240.png)
**Refactored 2017**
![Runtime 2017 Refactored Stock_Analysis](https://user-images.githubusercontent.com/104289098/168509672-77fe7afa-e6a8-4d91-8d80-fc028363a196.png)

**Original 2018**
![Runtime 2018 Original Stock_Analysis](https://user-images.githubusercontent.com/104289098/168509684-dbd58af3-bb1e-4c61-93ab-8c277514178e.png)
**Refactored 2018**
![Runtime 2018 Refactored Stock_Analysis](https://user-images.githubusercontent.com/104289098/168509700-4a11ce88-6c28-4eaf-ab9c-494f89fe9bab.png)


       
###  2017 vs 2018 Stock Analysis.

By Comparing the stock returns for 2017 and 2018 we can see that 2017 was a better year for stock in general with almost all stock having positive returns except for "TERP", the only two stock that had positive returns for 2018 were "ENPH" and "RUN". Steve's parents were interested in investing in "DQ" which had the highest returns on 2017 (199.4%) but reported losses in 2018 (-62.6%). By looking at the returns for the two years it looks as if it would be better for Steve's parents to invest in "ENPH".
 ![2017 Vs 2018 Stock_Analysis](https://user-images.githubusercontent.com/104289098/168509459-0705f848-7378-4ab6-b1a0-835450713c6e.png)
       
        

##  SUMMARY

###  Advantages and disadvantages of refactoring code in general.
    Refactoring the code can easily lead to errors or miss changes that need to be made with the new variables. It almost seems to be better to rewrite the whole macro from scratch to avoid "patching" although there will be cases where it might not just be possible to do because of the scope of the project. To make an analogy it seems as if refactoring code is like make an addition or a renovation to a house or building, the changes might be necessary but incorporating new things that were not contemplated in the original plan can make it challenging to make things work and it can create new difficulties that were not overseen.

  ###  Advantages and disadvantages of the original and refactored VBS script.
        The most obvious advantage of the refactored code is the reduction in run times which would allow to run an analysis of more stocks if needed be. The disadvantages of the reafactored code is the addition of lines of code and new variables used to create new arrays and set and index for those arrays. 

