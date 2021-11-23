# Stock Analysis with VBA Challenge

## Project Overview:

Steve, wants to analyze stock market data so that he can help his parents decicde which stocks are better for their portfolio. Steve has a dataset containing a list of green energy stocks that he should be able to analyze quickly and efficiently through VBA excel. 

### Project Challenge: 
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

## Project Results:

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


## Project Summary:

### **What are the advantages or disadvantages of refactoring code in general?**

Refactoring should not be thought of as just copying, pasting and reusing old code.  It can be very easy to mess up something in the refactoring process and create bugs or other problems that affect the functionality.  I found it a little less intimidating by going in manageable "chunk" rather than just doing the whole thing all at once.  Using comments is very helpful in keeping all the code organized and will down the line if anyone who might look at your code later.

**Advantages:**
- Flexibility - reusing code that you knows was written to solve one problem and applying it's logic to another problem. 
- Efficiency - Saves time in having to write the same code that was already written.
- Improve readability - having clear comments can make the QA and debugging process go more smoothly 
- Improve quality of the code so that future users can continue to build on code 
- Simplicity in that the originial code could be old and there may be new ways of simplifying that past logic

**Disadvantages:**
- Some code will need to be rearranged to function correctly. 
- Can't just copy and paste the code with out formulating a plan. This plan should be meant to help keep code blocks organized.
- Can be time conusming if need to conduct testing.  Adding the new code doesn't just go according to plan right away. 
- Not a catch all for bugs.  You Can mess up something in the process of refactoring and create new bugs or problems that affect the functionality. 

### **Adavantages and disadvantages of the original and refactored VBA script for this challenge?

Haing the original code saved me time from having to start over from scratch.  By reducing the amount the amount of code I had to write, it reduces the amount of error I could make.  The refactored code showed me how the logic could be simplified but putting it into action was another story. I ran into a lot of bugs and the performance improvements did not make much of a difference in processing time. 

**Advantages:**
- I did not have rewrite all the code that I created earlier in the module to get the challenge started. 
- The comments were very helpful in trying to determine the logic.

**Disadvantages:**
- I did experience bugs when trying to refactor.
- Some of the comments were vague and I could not figure out what I was trying to do. 
- It was hard to tell the difference between I did earlier in the module and what I did in the challange. 






































