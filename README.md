# week3-assignment-VBA-challenge
# Helping Steve’s parents analyze the stock market

## Background and purpose
#### In this project we are helping Steve’s parents to make an informed decision regarding which stocks have performed more successfully during the years of 2017-2018. Based on our analysis, Steve and his parents will be making a decision which stocks are better to invest in. We will be using Excel’s VBA tools and learning about manipulating data via defactoring a given code. We are interested in finding out which stocks have increased in their value and which stocks were traded more than others.

## Results
#### Having ran the analysis, we can conclude that 2017 and 2018 were drastically different for the majority of the stocks that we have data on. From the screenshots we can see that in 2017 almost all 12 stocks had positive return, except “TERP”. In 2018 10 out of 12 stocks had negative return. Here’s also an example of my code that helped me in my research:

###### If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value as well as If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value.

#### Having ran the original and defactored codes, we can spot a decrease in execution time down to 0.125 seconds. We can also see that the return on most stocks went down, even though some of their Total Daily Value has gone up between 2017 and 2018. 

## Summary

#### Refactoring code has its own advantages and disadvantages. A refactored code seems to work faster and appears more organized and understandable, with lots of comments. The main problem with defactoring is how time-consuming the process is. It almost felt like it would have been faster to build a completely fresh code, from ground up. What is also worth mentioning about the process of defactoring is debugging the original code, which is also very important.
