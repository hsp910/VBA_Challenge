# VBA Challenge

##Overview
The purpose of this project is to take VBA code made to analyze a set of data based on stocks and refactor it to become more efficient in runtime. The code is used to analyze the total daily volume and the return per year of various stocks being tracked. The code does this by matching stocks based on their tickers, small 2-4 letter series used to differentiate different stocks. The original issue at hand is that the code took much longer to run due to various issues involving the usage of for loops inefficiently. I came in with the purpose of streamlining the code so that the dataset would only have to be scanned once to make the code run much faster.
 
 ##Results
 For the sake of a baseline we must look at the original run times for both example sheets used in this project.
 ![2017 Example](Resources/VBA_Original_2017.png)
 This is the original code running through the 2017 stock data.
 ![2018 Example](Resources/VBA_Original_2018.png)
 This is the same code running through the 2018 stock data.
 Now that we understand the speed of the original code we can also understand how much more efficient the refactored code was.
 ![New 2017 Example](Resources/VBA_Challenge_2017.png)
 As we can see there is a large reduction in runtime from the original code to the refactored code for the 2017 sheet.
 ![New 2018 Example](Resources/VBA_Challenge_2018.png)
 Now we can also see that this is a repeated fact with the 2018 data sheet.
