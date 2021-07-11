# VBA_Challenge by Warren Camp
Title

Refactor VBA code and Measure Performance

I. Overview of Project

Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. 

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

II. Results

Significant execution time reductions were achieved through by using nested IF statements.  The nested IF statems reducted the number of execution conparsions and the number of index access to the 3 arrays.  
A total of 4 runs were made.  The runs were separated by data for 2017 and 2018 and by 2 sets of VBA code.  One set of VBA code was the challenge_starter_code code and the other was the update refactored code. 
Execution times for the 4 were for 2017: 0.8125 and for 2018: 0.7929688.  After refactoring the execurtion times were for 2017: 0.2617188 and for 2018: 0.265625. This represents execution being reduces by 68% for 2017 and 66.5% for 2018.

III. Summary

III. a. What are the advantages or disadvantages of refactoring code? 

The advantage is faster execution time, clearner more readable code and less bugs.  The disadvantage is longer development time and you it may not eb worth the effore if reduced execution is not achieved and if the system is always in a state of chnge.

III. b. How do these pros and cons apply to refactoring the original VBA script?

I spent about 50 % more time doing the refactoring.  I try many different things that did not work.  For example trying different methods to initialize the arrays and reduce the overhead associated with the Cells function and activate.  However, the great results in the reduction of execution time of 68% abd 66.5% clearly justify the extra time for recoding the code.
