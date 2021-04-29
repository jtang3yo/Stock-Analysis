# Stock-Analysis 

## Overview: VBA Stock Analysis Project
### Purpose
Purpose of this project is to refactor the VBA codes to loop in all data in different year one time to perform analysis on all stocks. First,we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
### Backround and challenge 
In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings. 
Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## Results: Refactor VBA codes and Search Performance 
### 1. Change the original set 2018 to yearValue 
* In order to loop in "2017" sheet, I refactored the original "2018" to "yearValue" 
* Range("A1").Value = "All Stocks (" + yearValue + ") "
### 2. The ticker is set equal to 0 before looping over rows
* Created a ticker variable and set it equal to zero before iterating over all the rows. 
* ticker![image](https://user-images.githubusercontent.com/82353749/116495396-ccc1f980-a870-11eb-8197-aad2537916f4.png)
### 3. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices 
* Created variables and declared them accordingly. 
* Created variables![image](https://user-images.githubusercontent.com/82353749/116495422-d9dee880-a870-11eb-8810-fcef61941434.png)
### 4. The ticker is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays, and loop through rows to retrieve ticker starting prices and ticker ending prices. 
* Looped over yearValue sheet, increased the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. 
* Loop over yearValue sheets![image](https://user-images.githubusercontent.com/82353749/116495444-e5321400-a870-11eb-940c-90bb9cf85e1d.png)
Stored values in tickerStartingPrices and tickerEndingPrices. 
### 5. Formatting the code 
* Made positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.
* Formatting![image](https://user-images.githubusercontent.com/82353749/116495466-f24f0300-a870-11eb-9d8a-6818a667a2dc.png)
### 6. Outputs for 2017 and 2018 stock analysis 
* Created button to run the current module 
* Generated analysis results for 2017 and 2018 
* 2017 results![image](https://user-images.githubusercontent.com/82353749/116495487-fda22e80-a870-11eb-8e34-5e2f892e7e73.png)
* 2018 results![image](https://user-images.githubusercontent.com/82353749/116495507-0430a600-a871-11eb-8e4f-46f500342a1a.png)
### 7. Final pop-up message showing the elapsed run time for the script
* VBA_Challenge_2017.png![image](https://user-images.githubusercontent.com/82353749/116494646-41943400-a86f-11eb-99b3-0b29b8fc970d.png)
* VBA_Challenge_2018.png![image](https://user-images.githubusercontent.com/82353749/116494655-4822ab80-a86f-11eb-95f6-5d92628cdfa7.png)

## Summary 
### Advantages of refactoring the codes: 
* Generally speaking, codes refactoring boosts the system performance and the refactored codes respond more quickly. 
* The codes run one time faster after being refactored, the original codes ran in 1.2 seconds to yield result in year of 2018, now it renders both 2017 and 2018 analysis results with almost same amount of time. 
* The codes allows user to access both 2017 and 2018 stocks by simply type the year in the input box, and render the analysis results, which is integrated and efficient. 
* Codes are fresher and easier to read after refactoring the variables, and less complex to execute with the button of "Run All Stock Analysis". 
### Disadvantages of refactoring the codes: 
* It takes more time to refactor the variables to enable to run searches on both 2017 and 2018 sheet, and there was significant amount of time spending on debugging and fixing errors when variables were not correcly declared, or architectural issues when refactoring some pieces of codes. 
* Defective or duplicate logical structures could affect the performance while or after refactoring. 
