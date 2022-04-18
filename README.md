# Stock-Analysis
## Overview

### Purpose: 
  The purpose of this project is to refactor an Excel VBA starter piece of code. This starter code assists Steve in knowing how fast VBA 
can compile results based on the year we plug in to find the respective stocks Total Daily Volume as well as its Return rate. 
The data we looked at consists of stock information for 12 different companies and with VBA, we compared the two years to see how well, 
or poorly, they did.


## Results
### Analysis of 2017
  The year of 2017 seemed to be a good year for almost all companies. With positive returns coming for 11 of the 12 companies. Even the 
company Steve’s parents wanted to invest in did well with a total daily volume of $35,796,200 with a 199.4% positive return rate. 
The only company with a negative return was TERP with a total daily volume of $139,402,800 and a negative return rate of  -7.2%. 
When running the code, I also noticed that the refactored time was faster than the original. The refactored time for the analysis 
ran in a time of about 0.68 seconds and the original code ran in a time of 0.71 seconds. Comparing the refactored and original, the refactored 
code ran faster. 

![All Stocks Analysis for 2017](./Resources/All_Stocks_Analysis_2017.pdf)

![Run code for 2017](./Resources/Code_Ran_2017.pdf)

### Analysis of 2018
  According to the table, in 2018, ten of the 12 companies all had negative return rates. What I also discovered was that the company 
Steve’s parents wanted to invest in actually did the worst out of all other companies. The total daily volume was $107,873,900 with a return 
rate of -62.6%. The two companies that did well (ENPH and RUN), had a positive return rate of 81.9% and 84% and a total daily volume of 
$607,473,500 and $502,757,100 respectively.  When comparing the code between the original analysis and the challenge, I noticed that the 
refactored code again ran faster than the original. The refactored code for 2018 ran in 0.68 seconds. Unlike the refactored code, 
the original code ran in 0.79 seconds. This shows that the refactored code for the challenge was faster than the original code for the Module.

![All Stocks Analysis for 2018](./Resources/All_Stocks_ Analysis_2018.pdf)

![Run code for 2018](./Resources/Code_Ran_2018.pdf)

## Summary 

### Advantages of Refactoring Code 
  When working with the refactored code, the first thing I noticed was the fact that the refactored code always ran faster than the original code. 
As stated previously, the refactored code times for 2017 and 2018 was around 0.68 seconds while the original code took about 0.71 seconds and 
0.79 for 2017 and 2018 respectively. In addition, the refactored code was more organized and easy to read, which can improve the look and 
quality of the code (such as the run times stated previously). In addition, using refactored code makes it easier to understand what code is 
doing. For example, when working on this challenge, there were comments on what had to be done to complete the challenge and I wrote my code 
under the comments. Working with the comments and having them explain what is going on also provides a form of organization for the whole code. 
With organized code, even someone who is not well versed in VBA can understand what is going on based on the comments and organization.

### Disadvantages of Refacotring Code 
  One issue when working with refactored code can be the sprout of errors that can occur if even one minute detail is missed. For example, 
when working with this refactored code, I had to make sure that the organized code ran correctly and was on the lookout for errors. 
When refactoring code, it is important to make sure all lines of code are properly organized and written or else there can be an error 
that can be time consuming to run through and check.

### Origonal code vs. Refacotred Code
  When a developer refactors code, it is done in order to restructure the code in a much simpler way. This can help if the original code is 
too messy from writing the original script. As stated before, when the code is messy and disorganized, it can be hard to read what exactly 
is going on. Despite changing the look of the code, the functionality stays the same. Refactoring even makes the code run faster. 
For example, the times for both 2017 and 2018 were the same but I believe that could be because of my hardware. The code ran faster than the 
original code by 0.2 to 0.3 seconds. This could be because of the simplicity of the refactored code causes less stress on my computer which makes 
it run faster as a result.
