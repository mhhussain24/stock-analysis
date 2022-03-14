# Analysis of Stocks
## OverView of the Project
### The purpose of the project is to analyze datasets including the entire stock market over any number of years. The code in the VBA attached loops through all the data provided, and then collects the information to quickly deliver the total daily volume and the return on each stock. The formatting of code will also allow for anyone viewing this code to quickly see what return percentages are below and above zero. 
## Results
### In the year 2017, the return on stocks was - for the most part - postive. Other than ticker "TERP" all the stocks had a positive return. Some stocks even had returns over 100%. However, in 2018 the return on stocks dropped substantially. There were only two stocks that were still maintaining positive returns. Of those two, one of them (RUN) actually increased its return percentage by about 80%. Total daily volume for 2017 was nearly 140 million less than in 2018 for the same tickers.
### The execution times of the original script and refactored script differed a substantial amount. The refactored script ran much faster. Please see image below for the macro run for year 2017 and then for 2018. As you can see, the run time has been cut down significantly. 
#### Embed 2017 Image
! [2017 Data](Resources/VBA_Challenge_2017.png)
! [2018 Data](Resources/VBA_Challenge_2018.png)
## Summary
### Refactoring code is beneficial as it makes code more efficient. It takes a written code and makes it so that new readers- people who have not been part of the process of when the code was initally written - to take a look at the code and have a decent understanding of what it is saying. Refactored code makes it easier to manipulate new data, since it is pretty nonspecific. It also helps to correct code that is initially imperfectly written. Not that refactoring perfects code entirely everytime, but it does tweak up the first draft for better use. 
### This applies to our refactoring of the original VBA Script too. For example, creating a tickerIndex to represent zero and then asking it to increase each time the next row's ticker does not match helps to create a code that allows for easy access to to the ticker array we set as well as the ticker Volume, ticker starting prices, and ticker ending prices. This one variable makes that process a lot simpler, especially when someone wants to add more tickers to analysis like this.