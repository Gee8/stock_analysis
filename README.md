# stock_analysis
## Stock analysis using VBA

### Overview of Project: Explain the purpose of this analysis.

The purpose of this analysis was to create a macro that will analyze stock data from years 2017 and 2018. The macro asks which year of data to look at, and returns a table with the each stocks ticker, total daily volume, the percent return for that year and returns a message box with the time it used to run the macro. We wrote two different versions of the macro. Initially, the macro from the module would have to loop through the entire sheet for each ticker to collect the data for each. For the challenge, we refactored our code from the module to find the data for each of our variables while only having to loop through the data sheets once.


### Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Shown below are the images of the results and run times for the 2017 module code, 2018 module code, 2017 refactored code, and 2018 refactored code, respectively. The stocks performed much better in 2017 than in 2018, and our results were the same with the different code. The major difference between our module code and refactored code was the run time. The refactored code achieved the same results as the module code in a fraction of the time. Although the code between each was similar, the major difference was creating the arrays:

`Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single`

by creating these arrays, we were able to increase our tickerVolumes, set the tickerStartingPrices and tickerEndingPrices by using the tickerIndex. For example, we used `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value` to increase the ticker volumes. This allowed us to gather the data required to create our table with one loop through the worksheets.

(inset images of data and run times here)


### Summary: In a summary statement, address the following questions.

1. What are the advantages or disadvantages of refactoring code?
- Some advantage of refactoring code include faster run times, more versatility in editing and applying the code in new scenarios, and condensed code which will be easier to edit for future use.
- A disadvantage of refactoring code would be the time required to refactor the original code. The time it takes to refactor may not be worth the effort for the advantages gained.

2. How do these pros and cons apply to refactoring the original VBA script?
- Our refactored code run in much less time, removed all hardcoded numbers and replaced them with an index. The refactored code could easily be changed to add additional stocks for the future by adding the new stocks to the tickerIndex. The original code would have to be edited more extensively to include additional stocks.