# Overview of Project
This analyst has been contacted by Steve, a recent finance graduate, to assist in the analysis of stock data for multiple green energy companies. Steve has selected 2017 and 2018 stock data for review and compiled it into a Microsoft Excel file. In addition to providing results of the analysis, this analyst has been tasked with generating reusable macros coded in Visual Basic for Applications (VBA)  for use in future analyses. An existing macro designed to perform a similar analysis was used as a baseline for the refactored macro. The new custom macro is designed to reliably generate consistent outputs for the twelve tickers selected by Steve.

# Results
The analysis was conducted by calculating the total daily volume and rate of return for the twelve green energy tickers selected by Steve.  Tickers with positive returns are highlighted in green and those with negative returns were highlighted in red. Results were generated for 2017 and 2018.

## 2017 Results
Eleven of the twelve green energy stocks experienced a positive return in 2017. The greatest return was seen in DQ with a rate of 199.4%. Only TERP had a negative return at -7.2%. (See image below)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/82056100/116831764-40207f80-ab7f-11eb-93ce-40417a561e80.PNG)

*2017 Green Stocks Results*

## 2018 Results
Ten of the twelve green energy stocks saw a negative return in 2018. The greatest loss was see in DQ at -62.6%. Only ENPH and RUN experienced a positive return that year with the latter seeing the greatest gain at 84%. (See image below)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/82056100/116831767-47e02400-ab7f-11eb-98b6-fb0f6a08367a.PNG)

*2018 Green Stocks Results*

## Code
Refactoring of the original subroutine proved beneficial in generating the results. A decrease in processing time was seen when generating results for each calendar year. This is due to the change made in the code requiring only one cycle through the dataset. The run time for the 2017 dataset dropped from .6171875 seconds to .09375 seconds. The run time for the 2018 dataset also saw a drop from .6015625 seconds to .0859375 seconds. (See images below)

![2017_Original](https://user-images.githubusercontent.com/82056100/116831889-14ea6000-ab80-11eb-8be3-3f06b9903ce1.PNG)

*Original Code - 2017 Dataset*

![2017_Refactor](https://user-images.githubusercontent.com/82056100/116831892-187de700-ab80-11eb-90dc-d501e1b25853.PNG)

*Refactored Code - 2017 Dataset*

![2018_Original](https://user-images.githubusercontent.com/82056100/116831894-1a47aa80-ab80-11eb-9fa2-c53a54863bad.PNG)

*Original Code - 2018 Dataset*

![2018_Refactor](https://user-images.githubusercontent.com/82056100/116831897-1caa0480-ab80-11eb-9c89-9aa2598412c4.PNG)

*Refactored Code - 2018 Dataset*

# Summary

## Refactoring Advantages and Disadvantages
The benefits of refactoring code include improved performance and providing a foundation for future improvements. Reviewing code like the subroutine provided for this effort is a necessity in volatile environments where inputs and requirements can change. Code must be updated to account for these fluctuations. On the downside, refactoring can be costly and financial impact should be considered before committing to code modifications. 

## Pros and Cons of This Refactored Code
Refactoring the original subroutine provided for this analysis resulted in quicker performance as discussed in the results section. This improvement will allow Steve to work with larger datasets in shorter time than with the subroutine???s predecessor. However, tailoring the code for this analysis has decreased the flexibility of the analysis as the data must be sorted by ticker and date in order to generate accurate results.


