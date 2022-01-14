#  Stocks Analysis

## Overview of Project
&nbsp;&nbsp;&nbsp; The purpose of this project was not only to analyze stock price increase or decrease from year to year but also to refactor code in VBA.  I tried different methods of coding before coming across the best outcome.  Here I will breakdown my inputs and explain in detail how everything works cleaner and more efficient.

### Purpose
&nbsp;&nbsp;&nbsp; The original code ran as intended for a small sample size.  I wanted to make sure the code could handle a larger data set and at optimal speeds.  I also needed to provide proof this code runs quicker than the original.

### Results
&nbsp;&nbsp;&nbsp; In this section I will layout my results in comparing stock prices from year to year and the code I refactored.  
Here is a snippet from the original code.  Take note of the nested for loop:
 <br /> 
 <br /> 
![original](https://github.com/LaszloCravensworth/stocks-analysis/blob/main/Original%20code.png)
 <br /> 
 <br />
&nbsp;&nbsp;&nbsp; Here is a look at the run times for 2017 and 2018 from the original code.  The orignal code for 2017 clocked a speed of 0.734375 seconds and 2018 at 0.75 seconds:
 <br />
 <br />
 ![origninalruntime2017](https://github.com/LaszloCravensworth/stocks-analysis/blob/d8dcbbccce3f82181af5750e2eb1f28f5ed8a1f1/Original%20Script%202017%20Runtime.png)
 <br />
 <br />
 ![originalruntime2018](https://github.com/LaszloCravensworth/stocks-analysis/blob/d8dcbbccce3f82181af5750e2eb1f28f5ed8a1f1/Original%20Script%202018%20Runtime.png)
 <br />
 <br />
 &nbsp;&nbsp;&nbsp; Now let's take a look at the refactored code.  Notice the addition of a new variable named tickerIndex that will store the array before looping.  I also removed the nested for loop and added a separate for loop at the end of the code.  In doing all of this it made the code cleaner and quicker.  Take a look at the refactored code and see the differences from the original code.
 <br />
 <br />
 ![refactoredcode](https://github.com/LaszloCravensworth/stocks-analysis/blob/d8dcbbccce3f82181af5750e2eb1f28f5ed8a1f1/Refactored%20Code.png)
 <br />
 <br />
 &nbsp;&nbsp;&nbsp; In these next two screenshots you can clearly see the code ran up to 6 times faster than the original.  The speed for the 2017 refactored code came in at 0.1171875 and the 2018 at 0.125.  The purpose of this project was completed.
 <br />
 <br />
 ![refactored2017](https://github.com/LaszloCravensworth/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)
 ![refactored2018](https://github.com/LaszloCravensworth/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)
 <br />
 <br />
 &nbsp;&nbsp;&nbsp; As you can see, all but two stocks decreased in value from this sample size from 2017 to 2018.   You would have been giving back some of your returns from 2017 in 2018 in everything except $ENPH and $RUN.  The stock market can have volatile movements in one direction or the other.  This stock analysis clearly shows what could happen if you chose the "wrong" stock.
 <br />
 <br />
 &nbsp;&nbsp;&nbsp; In conclusion, this project was a great experience refactoring code.  Now let's shed a little light on the positives and possible drawdowns of refactoring.  I could see this made my original code look slow and unefficient.  This could turn a mediocre program into an amazing one.  It could run faster.  It would have more clarity which would make it easier to understand for future edits. A couple drawdowns for real world applications would be time and money.  You could run out of money and then time refactoring a program that worked. I look forward to more projects editing code to make it cleaner and more efficient.
