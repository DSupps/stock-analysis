# Stock Analysis with VBA Challenge

## Section 1: Project Overview

Steve, wants to analyze stock market data so that he can help his parents decicde which stocks are better for their portfolio. Steve has a dataset containing a list of green energy stocks that he should be able to analyze quickly and efficiently through VBA excel. 

## Purpose 
*The purpose of this project is to help Steve analysis stock data quickly and efficiently by adding some new functionality to previously used code  By refactoring this code in VBA, I want to see if I can make it more efficient - by taking fewer steps, using less memory, or improving the logic of the code.  In addition, I want to make the code easy to read just in case somebody else needs to make updates.*

## Analysis and Challenges
The Analysis was performed by refractoring Module2_VBA_Script so that when the code loops through the data one time and collects all the information needed. The new refractored code should run faster than it did prior to be being refractored. 

### Analysis Requirements

- Download and prepare dataset and rename it VBA_Challenge.vbs
- Create a folder called resources to hold the run-time pop-up messages that you’ll screenshot after running refactored analyses for 2017 and 2018
- Rename the green_stocks.xlsm file that you used in this module as VBA_Challenge.xlsm
- Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor
- Add code to refractored VBA where indicated
- There are comments to explain the purpose of the code
- Outputs for the 2017 and 2018 stock analyses
- Pop-up messages showing the elapsed run time for script

## Section 2: Results

**1a) Create a ticker Index variable and set it equal to zero before iterating over all rows**
![1a](https://user-images.githubusercontent.com/36451701/116830468-570fa380-ab78-11eb-809f-01c711581e3a.png)

**1b) Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices**
![1b](https://user-images.githubusercontent.com/36451701/116830575-43187180-ab79-11eb-94db-b9cc3e3f7786.png)

**2a) Create a for loop to initialize the tickerVolumes to zero**![Uploading 2a.png…]**
![2a](https://user-images.githubusercontent.com/36451701/116830616-807cff00-ab79-11eb-9144-ae5b12f24f9f.png)

**2b) Create a for loop that will loop over all the rows in the spreadsheet**
![2b](https://user-images.githubusercontent.com/36451701/116830895-698adc80-ab7a-11eb-99da-923d15a26190.png)

**3a) write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker**
![3a](https://user-images.githubusercontent.com/36451701/116830923-8c1cf580-ab7a-11eb-9168-c17116c30f0f.png)

**3b) Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable**
![image](https://user-images.githubusercontent.com/36451701/116830942-a7880080-ab7a-11eb-807e-764b57cce9e3.png)

**3c) Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable**
![3c](https://user-images.githubusercontent.com/36451701/116830978-d00ffa80-ab7a-11eb-9eea-3a3966f7e568.png)

**3d) Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker**
![4](https://user-images.githubusercontent.com/36451701/116831067-401e8080-ab7b-11eb-9439-36b42ff0bffe.png)

**4) Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet**
![Step 4](https://user-images.githubusercontent.com/36451701/116831184-cd61d500-ab7b-11eb-9d5d-38c404db1fb7.png)

**5) Run the stock analysis, then confirm that your stock analysis outputs for 2017 and 2018 are the same as they were in the module**

### 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/36451701/116831624-89240400-ab7e-11eb-811c-0c28129eea9e.png)

### 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/36451701/116831647-9e009780-ab7e-11eb-98ea-80c4c19fbb95.png)


## Section 3: Summary





































