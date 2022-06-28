# Stock Analysis

## Overview of Project
Client Steve wants to do more research for his parents related to the best stock in which to invest. He was pleased with the analysis done on the provided dataset, however now he wants to expand the dataset to include the entire stock market over the last few years. Given this new task, we will need to refactor our code to be able to run the the stock analysis code faster, and for it to be able to include more data in the futre when provided with larger datasets.

### Purpose 
 The purpose of this challenge was to determine whether refactoring previously written code successfully makes the VBA script run faster and provides the same data analysis result. Once results are produced, data will be analyzed in order to help Steve make his decision, about which stocks would be best for his parents to invest in given the 2017 and 2018 trends.

## Analysis

### Analysis 2017
When analyzing all stocks in the data set in 2017, using the refactored code, it appears all but one stock, "TERP" , provided profitable returns. Amoung the profitable stocks there were four stocks that returned over 100%. Those stocks were "DQ", "ENPH", "FSLR" and "SEDG", with the highest performing stock being "DQ", see *Figure 1* below. Now, in our previous anlysis of year over year, comparing 2017 to 2018, it was noted that the "DQ" stock did not perform as well in 2018, and in fact had a significant decrease in return, see *Figure 2* below in Analysis 2018. Also note *Figures 1a-b*, the runtime for the previous original code in *Figure 1a*, and the refactored code runtime in *Figure 1b*. It is clear that the refactored code is more efficient , by taking less time to run.

#### Refactored 2017 All Stock Analysis Data

*Figure 1* 

![VBA_Challenge_2017](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2017.png)

*Figure 1a*-Original Code

![VBA_Challenge_2017_Runtime_NonRefactored](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2017_Runtime_NonRefactored.png)

*Figure 1b*-Refactored Code

![VBA_Challenge_2017RuntimeRefactor](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2017RuntimeRefactor.png)



### Analysis 2018
 When analyzing all stocks in the data set in 2018, using the refactored code, it appears all but two stocks, "ENPH" and "RUN" , provided non-profitable returns. Amoung the non-profitable stocks it is important to note that three out of the four profitable stocks in 2017 were no longer profitable in 2018, including "DQ", "FSLR" and "SEDG". It is also important to note that  in 2017, "RUN" was not profitable, however in 2018 it provided an 84% return. The only consistently profitable stock across both years is "ENPH", see *Figure 1* above and *Figure 2* below. Also note *Figures 2a-b*, the runtime for the previous non-refactored code in *Figure 2a*, and the refactored code runtime in *Figure 2b*. It is clear that the refactored code is more efficient , by taking less time to run.

## Refactored 2018 All Stock Analysis Data

*Figure 2*

![VBA_Challenge_2018](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2018.png)

*Figure 2a*-Original Code

![VBA_Challenge_2018_Runtime_NonRefactored](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2018_Runtime_NonRefactored.png)

*Figure 2b*-Refactored Code

![VBA_Challenge_2018RuntimeRefactor](https://github.com/maureengamache/Stock-Analysis/blob/main/VBA_Challenge_2018RuntimeRefactor.png)
 

## Results
In conclusion, the best stock for Steve's parents to invest in given the current dataset would be stock "ENPH". "ENPH" is the best stock to invest in as it was profitable across both years, 2017 and 2018. Also, when runinng the refactored code it was noted that both the 2017 and the 2018 analysis took less time to run (i.e. were faster in producing the output analysis) by approximately 0.6 seconds. Therefore, it is a clear advantage that faster runtimes are produced when refactoring original code. However, a disadvantage of refactoring code is the concept of "re-creating the wheel". Meaining that the developer has to find a new way to do the same thing that the orginal code did, but more efficiently. The pros of refactoring the original VBA code are that more data can be included in the future when the refactored code is used, in addition to fast run times. The cons of refactoring the original VBA code are the time commitment and the extensive knowledge of VBA required to create and Debug anything in the refactored code that is not working properly.   


