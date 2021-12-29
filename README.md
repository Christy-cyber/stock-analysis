# Analysis of Green Stocks

## Overview of Project

Steve investigated the performance of DQ stock, a green stock, for 2018 to determine if it would be a good investment for his parents.  He found that it did not perform well in 2018 (-62.6% return) so inquired about the performance of other green stocks that would be good potential investments.  Steve analyzed 12 green stocks, including DQ, from 2018 and found that two, ENPH and RUN, had positive returns (81.9% and 84.0%, respectively).

### Purpose

The purpose of this exercise was to
  1) further investigate the return on the same 12 green stocks in 2017
  2) refactor VBA code that would result in faster processing of data output for these 12 stocks for 2017 and 2018

## Results
### Performance of Green Stocks in 2017 Compared to 2018
Eleven of 12 green stocks had positive returns in 2017, with four having over 100% returns (DQ, ENPH, FSLR, and SEDG).  Only one stock, TERP, showed a negative return (-7.2%).

![VBA_Challenge_2017--Stock Performance Only](https://user-images.githubusercontent.com/95387273/147701968-3d2acc71-6e47-4447-b046-092668f5ecb7.png)

Only two of the same twelve stocks, ENPH and RUN, had positive returns in 2018 (81.9% and 84.0%, respectively).

![VBA_Challenge_2018--Stock Performance Only](https://user-images.githubusercontent.com/95387273/147702207-8fc37f9d-7173-433a-a0f7-d090814ce0f3.png)

### Performance of Refactored VBA Code


## Summary
### Advantages and Disadvantages of Refactoring Code in General
Refactoring code can save a lot of time because not all of the code has to be rewritten.  The most useful sections can be copied and pasted, and, in some cases, only a few lines may have to be added or deleted.  One of the disadvantages is that, in some cases, it is easier to write the code from scratch than try to follow someone else's logic, especially when notations have not been included in the code. Debugging can also be an issue and may require more time, especially when significant changes to the original code are being made.

### Advantages and Disadvantages of the Original and Refactored VBA Script
The original VBA script worked well on a small number of stocks for both 2017 and 2018.  Writing code and initializing an array of 12 ticker symbols did not take much time and the output was provided in less than half of a second.  The advantage of the refactored script, though, was that the time to analyze the data was much faster and final output was provided in less than a tenth of a second.  This may not seem like so much of a difference for 12 stocks as one waits for output, but as the number of stocks increases in the analysis, the wait time would also increase and be noticeable.  A major disadvantage of the original code is that it would not work well for a large data set.  It did not take much time to initialize an array for 12 tickers, but one would not want to do this for many ticker symbols.  Why would someone want to initialize 100 ticker symbols if they did not have to?  It would be easier to rewrite the code to automatically harvest ticker symbols as part of the output.  This was also the same disadvantage that I have for the refactored VBA script.  While the overall time to create output decreased with the refactored script, the code would not be useful for a large data set and would need to be refactored to harvest ticker symbols automatically.

### Recommendations for Selecting Green Stocks
Steve should recommend the green stocks ENPH and RUN to his parents because of their high return rates over the course of two years.
