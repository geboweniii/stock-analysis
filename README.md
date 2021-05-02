# Overview of Project
This analyst has been contacted by Steve, a recent finance graduate, to assist in the analysis of stock data for multiple green energy companies. Steve has selected 2017 and 2018 stock data for review and compiled it into a Microsoft Excel file. In addition to providing results of the analysis, this analyst has been tasked with generating reusable macros coded in Visual Basic for Applications (VBA)  for use in future analyses. An existing macro designed to perform a similar analysis was used as a baseline for the refactored macro. The new custom macro is designed to reliably generate consistent outputs for the twelve tickers selected by Steve.

# Results
The analysis was conducted by calculating the total daily volume and rate of return for the twelve green energy tickers selected by Steve.  Tickers with positive returns are highlighted in green and those with negative returns were highlighted in red. Results were generated for 2017 and 2018.

## 2017 Results
Eleven of the twelve green energy stocks experienced a positive return in 2017. The greatest return was seen in DQ with a rate of 199.4%. Only TERP had a negative return at -7.2%. (See image below)

## 2018 Results
Ten of the twelve green energy stocks saw a negative return in 2018. The greatest loss was see in DQ at -62.6%. Only ENPH and RUN experienced a positive return that year with the latter seeing the greatest gain at 84%. (See image below)

## Code
Refactoring of the original subroutine proved beneficial in generating the results. A decrease in processing time was seen when generating results for each calendar year. This is due to the change made in the code requiring only one cycle through the dataset. 

# Summary

## Refactoring Advantages and Disadvantages
The benefits of refactoring code include improved performance and providing a foundation for future improvements. Reviewing code like the subroutine provided for this effort is a necessity in volatile environments where inputs and requirements can change. Code must be updated to account for these fluctuations. On the downside, refactoring can be costly and financial impact should be considered before committing to code modifications. 

## Pros and Cons of This Refactored Code
Refactoring the original subroutine provided for this analysis resulted in quicker performance as discussed in the results section. This improvement will allow Steve to work with larger datasets in shorter time than with the subroutine’s predecessor. However, tailoring the code for this analysis has decreased the flexibility of the analysis as the data must be sorted by ticker and date in order to generate accurate results.
