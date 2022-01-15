# Analysis of Stock Volumes and Returns

## Overview of Project: 

### The purpose of this project was to refactor code in VBA which looped through stock data to find the Total Daily Volume and Return for a list of twelve different stocks. The purpose of refactoring the code was to run it so that it only had to loop through the data once, as opposed to running it for each individual stock, so that the client could analyze data for thousands of stocks with greater effeciency. 

## Results:

### Stock Peformance 

After performing the analysis on the twelve stocks for the years 2017 and 2018, it is apparent that the returns were greater in 2017 with DQs return reaching 199.4% while dropping -62.6% in 2018. In fact, most of the stocks, with the exception of RUN which rose to 78.5% in 2018, dropped to a negative return between the years of 2017 and 2018. When comparing the Total Daily Volumes for each year, we can see that the number of shares increased for 58% of the stocks between 2017 and 2018 even though returns decreased for 83% of the analyzed stocks. 

![AllStocks_2017](https://user-images.githubusercontent.com/95589611/149638383-75068dfb-78d5-4272-bef0-a5b830212664.png)
![AllStocks_2018](https://user-images.githubusercontent.com/95589611/149638387-7c46d74a-8968-4433-b2d7-6aa8a8e60daf.png)

### Execution Times Between Original and Refactored Scripts

Refactoring the original script greatly improved the effeciency of our code. For 2017, our refactored code decreased the execution time by 2.33 seconds (86% decrease); For 2018 there was a decrease in execution time by 2.23 seconds (85.7% decrease).

**Original Time (2017):**
![Original_Time](https://user-images.githubusercontent.com/95589611/149638413-405eebae-2017-47f4-b86f-955b04cc3669.png)

**Refactored Time(2018):**
![RefactoredTime_2017](https://user-images.githubusercontent.com/95589611/149638417-b26c1edf-5912-419b-8a03-2f4c93108f68.png)

**Original Time (2018):**
![Original_Time_2018](https://user-images.githubusercontent.com/95589611/149638423-f1e547f8-60ff-443c-b561-69f09f9ef4be.png)

**Refactored Time (2018):**
![RefactoredTime_2018](https://user-images.githubusercontent.com/95589611/149638431-754cb652-0547-4653-812b-b802fa5373f1.png)

The refactored script decreased the execution time by isolating the values from the tickers so that the output was simply a list of values which made the code run faster because we only had to loop through the data once and the Macro was not concerned with connecting the values to their tickers. 
In the original code, we created an array to hold twelve stocks and assigned those stocks an element in the array. 
![Original_1](https://user-images.githubusercontent.com/95589611/149638498-2754d750-6b30-4ef7-849d-60182da34263.png)

We looped through each element (tickers) to find the correlating values of Starting Price, Ending Price, and Volumes. 
![Original_Loop](https://user-images.githubusercontent.com/95589611/149638903-10727e89-af5e-410e-9475-0f8211bd3445.png)

In the refactored code we stored the values for Starting Price, Ending Price, and Volumes in their own arrays which held the values for all of the tickers. I then looped through these arrays to assign the output to their correlating range of cells. 
![Refactored_Loop](https://user-images.githubusercontent.com/95589611/149639067-d826ea93-2363-44e1-96c2-55f4a6baf4a6.png)
![Refactored_output_loop](https://user-images.githubusercontent.com/95589611/149639153-e6738e77-25bd-4c22-88e9-d7e201aa3a0d.png)

## Summary 

### What are the advantages or disadvantages of refactoring code?

**Advatages:** The advantages of refactoring a code are that it can decrease execution time significantly as we have seen with the example of refactoring our script for the Stock Anaysis. Shortening the execution time comes with its own advantages, such as the advantage displayed by our current project in which our client can now run through data from thousands of stocks in a short amount of time. Another advantage is that it can make a code cleaner and shorter which may make replicating or producing the code easier and faster for the analyst. 
**Disadvantages:** A disadvantage of refactoring code is that it can make the code harder to interperate for anyone hoping to understand the refactored code's function. A refactored code has the goal of making the macro run faster, so it is formatted in a way that is geared toward aiding the software rather than the analyst who is deciphering the code. Another disadvantage of refactoring a code is that the analyst must make the decision whether or not taking the time to refactor the code is worth the trouble when weighed against how much time a refactored code is going to save the analyst when executed. 
### How do these pros and cons apply to refactoring the original VBA script?

In the context of the current project, refactoring the VBA stock analysis script decreased our execution time by 86%. This will allow our client to analyze thousands of stock data without having a script that takes a long time to run. This shorter script also allows the client to replicate and modify the new script without having to read through many different loops and lines of code. The disadvantage of the refactored code is that its summarized state makes it harder to understand. It also means that there had to be time put into figuring out how to refactor the original code so that it ran. 
