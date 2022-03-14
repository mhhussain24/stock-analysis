# Analysis of Stocks
## OverView of the Project
### The purpose of the project is to analyze datasets including the entire stock market for any year that data is provided for. The code in the VBA attached loops through all the data provided, and then collects the information to quickly deliver the total daily volume and the return on each stock. The formatting of code will also allow for anyone viewing this code to quickly see what return percentages are below and above zero. 
## Results
### Comparing 2017 and 2018
#### In the year 2017, the return on stocks was - for the most part - postive. Other than ticker "TERP" all the stocks had a positive return. Some stocks even had returns over 100%. However, in 2018 the return on stocks dropped substantially. There were only two stocks that were still maintaining positive returns. Of those two, one of them (RUN) actually increased its return percentage by about 80%. Total daily volume for 2017 was nearly 140 million less than in 2018 for the same tickers. 
### Code to Describe Analysis
#### To find the total daily volume we first set tickerVolumes as "Long." Then, we created a for loop to initialize volume to zero. We did this by setting "i" equal to "0 to 11" and then "tickerVolumes(i) = 0." We created another for loop to run through all the rows and increase the volume for each ticker by using code   "tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value." For the return, we use the formula "tickerEndingPrices(i) / tickerStartingPrices(i) - 1" for each ticker. In order to ensure we have the first and last row's of the selected ticker we want for each formula, we use "If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value End If" and  "If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value." The code runs asynchronously, so after each run the ticker index increases by one, and the code is run again for the next ticker.  
### Run Time Comparison and Screenshots
#### The execution times of the original script and refactored script differed a substantial amount. The refactored script ran much faster. Please see image below for the macro run for year 2017 and then for 2018. As you can see, the run time has been cut down significantly. 
##### Embed 2017 Image
! [2017 Data](Resources/VBA_Challenge_2017.png)
! [2018 Data](Resources/VBA_Challenge_2018.png)
## Summary
### Benefits (Generally)
#### Refactoring code is beneficial as it makes code more efficient. It takes a written code and makes it so that new readers- people who have not been part of the process of when the code was initally written - to take a look at the code and have a decent understanding of what it is saying. Refactored code makes it easier to manipulate new data, since it is pretty nonspecific. It also helps to correct code that is initially imperfectly written. Not that refactoring perfects code entirely everytime, but it does tweak up the first draft for better use.
### Disadvantages (Generally) 
#### Refactoring code is a time consuming process. There is also a chance that you may mess up your original code while trying to refactor it for effeciency purposes. You may lose track of what you have oversimplified. IF you have messed something up half way through refactoring, it could be quite time consuming again to figure out where you went wrong. 
### Benefits for this VBA Script
#### This applies to our refactoring of the original VBA Script too. For example, creating a tickerIndex to represent zero and then asking it to increase each time the next row's ticker does not match helps to create a code that allows for easy access to to the ticker array we set as well as the ticker Volume, ticker starting prices, and ticker ending prices. This one variable makes that process a lot simpler, especially when someone wants to add more tickers to analysis like this.
### Disadvantages for this VBA Script
#### Refactoring this VBA script did take a significant amount of time. There were also places where I felt I had done essentially the same steps but really messed up and was a bit frustrated trying to figure out where I went wrong from the original script. I spent quite a bit of time having to retest if my code was providing the same functionality after I had tweaked it. 
