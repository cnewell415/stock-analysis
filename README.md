# VBA of Wall Street

## Overview of Project
Refactor our workbook for Steve so that it runs faster for future scalability.  
Our goal is to loop through all the data one time and still collect all the required stock information. If we can do this, does the code run any faster? If so, why? 

## Analysis and Challenges
### Investment Recommendation
![2017 Results](/2017_Results.png)![2018 Results](/2018_Results.png)
When you compare the 2017 and 2018 stock analysis results side by side you can visually see that 2017 was generally a more successful year for all companies. The majority of the companies were unable to repeat their 2017 success and posted negative returns. 

Two companies go against that trend, ENPH and RUN. Both of them were able to post positive returns in back-to-back years. 

Limited to this data set, I would recommend Steve look closer into ENPH and RUN as potential companies to diversify his parent's investments in.

### Code Runtime Results
When it comes to timing measurements, it is good practice to measure numerous times and calculate the average. I ran four sets of measurements in this fashion. 

1) AllStocksAnalysis() | Input = 2017
2) AllStocksAnalysisRefactored() | Input = 2017
3) AllStocksAnalysis() | Input = 2018
4) AllStocksAnalysisRefactored() | Input = 2018

I ran each of these 5 times, then calculated the average runtime for each. Then I compared the runtime of the refactored macro to the original macro. The results for both years shows that the refactored code runs quicker. 

![Runtime Results](/Runtime_Data.png)

The code runs faster after the for loops in the original code are separated and do not run within one another. 

**Original Code**
```
For i = 0 To 11
   For j = 2 to 3013
   Next j
Next i
```

The inner j loop runs 3011 times ever time the outer i loop iterates. Therefore, these nested loops will run 36,132 iterations for the data sets we have. 

**Refactored Code**
```
For i = 0 to 11
Next i

For j = 2 to 3013
Next j

For i = 0 to 11
Next i
```

After reworking the logic to enable three separate For loops drastically decreases the number of iterations to complete the code. The new refactored code will run the first loop 12 times, the second loop 3,011 times, and the third loop 12 times for a total of 3,035 iterations. Roughly, 33,000 fewer runs through the code, thus a faster runtime. 

## Summary

Refactoring code results in faster runtimes by efficiently using memory and resources. This is helpful to create code that can scale upwards to work on larger datasets while not bogging down the machine. 

A possible disadvantage of refactoring code would be code that is too simplified. Is the code too streamlined for another programmer to understand it? Would that be useful in a company setting? Most likely not, since employees are always changing. 

In the case of our AllStocksAnalysis code, refactoring was useful in speeding it up and should have been done. 

