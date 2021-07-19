## VBA of Wall Street - Written Analysis of the Results

### Overview of Project: Explain the purpose of this analysis.

  The purpose of this analysis was to edit or refactor the VBA code to loop through all the stock data one time in order to collect an entire dataset and to determine if after refactoring, the code made the VBA script run faster by measuring the performance. 

### Results: 

![Stocks_2017](https://user-images.githubusercontent.com/86751774/126190534-54333b2c-cb73-4d09-973a-2722a6d45604.png)

![Stocks_2018](https://user-images.githubusercontent.com/86751774/126190553-48c02d9c-3f27-41fe-bb3e-c53b28c7cb20.png)

  From what is displayed above, the stock performance in 2017 was much better than the stock performance in 2018. In 2017, only one stock out of the 12 had a negative return whereas in 2018, all but 2 out of 12 stocks had a positive return.

   Elapsed run-time for the original code:

![green_stocks_2017](https://user-images.githubusercontent.com/86751774/125967040-6be9fb7a-e7a6-4055-8ab2-1b3891297e0e.png)

![green_stocks_2018](https://user-images.githubusercontent.com/86751774/125968483-8e261ddb-d948-4bfb-a36d-b729b2ae2b2b.png)

   Elapsed run-time for the refactored code:
   
  ![VBA_Challenge_2017](https://user-images.githubusercontent.com/86751774/125832159-a3dd53fa-704b-41ac-a1cc-6bf01694030b.png)
  
  ![VBA_Challenge_2018](https://user-images.githubusercontent.com/86751774/125832700-85eb655d-9922-4511-8a02-7092b815f886.png)
  
  The refactored code definitely ran faster than the original code and that is shown in the above images. The refactored code for year 2017 ran 0.097656 times faster than the original code. The refactored code for year 2018 ran 0.183593 faster than the original code. 
  
### Summary:

  #### What are the advantages or disadvantages of refactoring code?
  
   The major advantage of refactoring code is being able to make the original code more efficient than before. However, the disadvantage in refactoring a code that already works is that we might be making it deviate from a full working code to something that won't work as well or won't work at all. 
   
  #### How do these pros and cons apply to refactoring the original VBA script?
   From the comparing of the elapsed run-time for the original and refactored code, it can be concluded that refactoring made the script run faster than before. In terms of the disadvantage in refactoring, as the script was much longer than what I was used to, a mistake in any lines of the code and the code would not run as it should. 
