# VBA Challenge on Stock Analysis

**Purpose**
The purpose of this project was to take existing data and code within VBA and an a “refractor” element/script to our VBA macro as a means of looping through all data to collect the entire dataset in one command. From this, we will compare our refractor code’s total affect on the run time of our VBA script.
Overview and Method
A tickerIndex was created and set to 0 before all rows were looped through.

![image](https://user-images.githubusercontent.com/91284661/136671367-17ee7c13-be6e-459f-88ad-375d75a37584.png)
 
From this, new arrays were generated for tickers i.e. tickerVolumes, tickerStartingPrices, and tickerEndingPrices. All should be Single with the exception of tickerVolumes which was Long instead of Single. Following this, a nested For Loop was generated which effectively used tickerIndex to access the stock ticker index for tickerVolumes, tickerStartingPrices, and tickerEndingPrices respective arrays. One loop was created to initialize tickerVolumes (set to zero) and  see if the next ticker in the row matched, if not, it would increase the tickerIndex.
Next, a For loop was created to initialize the tickerVolumes and set it to zero. If the next row’s ticker doesn’t match our current row, increase the tickerIndex. 

![image](https://user-images.githubusercontent.com/91284661/136671395-a79d49ee-00fe-403a-a835-677344bfa25a.png)

The script then read through all stock data and stored the values from respective ticker, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. From this, the scrip that increased tickerVolumes and adds the ticker volume for current stock ticker.

![image](https://user-images.githubusercontent.com/91284661/136671403-6940ab42-6e13-4880-82ec-a5daf4cd3670.png)

From this, an if-then statement checked current rows was the first row with the selected tickerIndex. Assuming so, it assigned current closing prices to tickerStartingPrices and tickerEndingPrices.

![image](https://user-images.githubusercontent.com/91284661/136671406-4ab932eb-09e2-4233-9625-a91a00395021.png)

With our script written, we then went and actually performed our Stock Analysis for years 2017 and 2018, images of which should be applied on GitHub and shown below. 

**2017**
![image](https://user-images.githubusercontent.com/91284661/136671413-bb6c3eff-63dc-45bb-94fe-1083b7433fb8.png)

**2018**
![image](https://user-images.githubusercontent.com/91284661/136671466-50de3cd0-4a82-4a50-ab17-5733557738d6.png)
 
Also included as part of our script and analysis was the usage of a timer to help show whether our code was indeed improving efficiency of analyzing our data set. The refracted for years 2017 (.84375 seconds) and 2018 (.8554688 seconds) are shown below. 

**2017**
![image](https://user-images.githubusercontent.com/91284661/136671474-75bce83d-6037-4da7-8b5f-0bbdcabf4da7.png)

**2018**
![image](https://user-images.githubusercontent.com/91284661/136671418-3544375d-b5f7-4e76-a631-2d03717f3574.png)
 
**SUMMARY**
Some **advantages** of the usage of refactoring code are, in no particular order:
	Logical errors are more easily shown in appearance when nested loops are used.
	VBA through Excel visualization and interpretation of code shows patterns that are more easily seen.
Some **disadvantages** of the usage of refactoring code are:
	Possibility of redundant or duplicated lines of code throughout your script
	Too complicated a code may not be best tackled through refactored code, instead, it may be better to split your goals and tasks into several discrete functions.
	Possibility of the refactoring process affecting your outcomes.

In our case, the pros and cons can certainly be seen in our original VBA script through the readings of the timer index for both the non-refactored and refactored code respectively. All in all, our refactored code DID indeed make our script run faster. In the case of testing for both years, the script ran in almost half the time. Admittedly, though only performed on my computer and the results being almost a second off the computing time for both years (for a total of almost two seconds per the images below). Still, the insight seems clear. The larger the data pool you are working with, the more of a benefit on pure run-time-efficiency refactoring your code becomes.

**Non-Refactored 2017**
![image](https://user-images.githubusercontent.com/91284661/136671430-e3540cce-b42b-491c-9aa8-761590b22993.png)

**Non-Refactored 2018**
![image](https://user-images.githubusercontent.com/91284661/136671434-09772ba6-9aa7-4405-a596-5948a3af013a.png)

