# Stock Analysis

## Overview of project

### The purpose of this project is to help a client automate the analysis of stockmarket data using Visual Basic for Applications. VBA code was written including a nested `For` loop that analysed a small data set. The code was then refactored to run more efficiently with a single loop in order to analyze larger stock datasets.

## Results

### Original VBA Code from Module 2

In the original code, a nested `For` loop was created to run through the entire dataset from top to bottom multiple times before returning the analysis output. When running through the data this way it appeared to be an instantaneous result, but adding a timer allows us to see precisely how long the code ran. Below is attached the timer results from the 2017 and 2018 stock market dataset analyses from the workbook.

![Original_Code_2017](https://user-images.githubusercontent.com/111471057/188507732-a3e1eea2-ab3c-4a42-8459-ceb2ee98719c.png)       ![Original_Code_2018](https://user-images.githubusercontent.com/111471057/188507738-24a07831-976a-45a0-ace5-502d98e8e692.png)

### Refactored VBA Code from Challenge

For the challenge, the VBA code has been edited, or refactored, to run more efficiently by removing the nested `For` loop and using an added index variable. With this adjustment the code will run through the entire dataset once before returning the analysis output. The same timer is included in the refactored code to show the improved run time, shown below. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/111471057/188508027-d1da57a8-b423-44fe-addd-6b1eaef71f40.png)     ![VBA_Challenge_2018](https://user-images.githubusercontent.com/111471057/188508030-4218a0fe-7343-4ce7-a1da-4ada127a3280.png)


### Summary

One of the advantages to refactoring code is to increase the speed or efficiency of the run time as well as reducing the strain put onto the computer memory. No one likes to wait on their computer to process a request! One of the disadvantages of refactoring is that it is time consuming. The example code was fairly simple, but the time spent on writing it increased in order to decrease the run time. Another advantage to refactoring, however is increased readability of the code. Even though it took more time to refactor the code, it will take less time in the future to review, update, or add to the code.

Looking at the original VBA script, one of the advantages is that the data used in the workbook is not very large. The original script still pulls the data in less than a half of a second. One of the disadvantages is that the script cannot easily be used for a larger dataset because it will increase the amount of memory needed to provide the analysis output. Aside from the actual run time, both the original code and the refactored code had the same output results, however the advantage of the refactored code is that it can be reused on much larger datasets because the refactored code will run faster and use less computer memory to show the analysis. 
