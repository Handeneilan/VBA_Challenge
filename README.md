# VBA_Challenge - DQ Analysis

## Overview of Project

The purpose of this analysis is to create a macro which will help our client Steve to analyze stocks for his parents. This tool will help him identify what the returns on investements are and how he can help his parents invest. We are using the data from 2017 and 2018 to make a good prediction.

## Results

Analysis was done by using Excel VBA, and two tables were created to help with the Client’s question.

![2017](https://user-images.githubusercontent.com/104239978/183552485-397f9005-54b6-4682-9984-cb567429f921.png)

![2018](https://user-images.githubusercontent.com/104239978/183552498-14b5a0bf-57da-4068-813f-772280e8347e.png)

It looks like in 2017 is very safe to invest in almost any of the analyzed stocks. Only  (TERP) had a negative return, and 6 of them has returns over %50. DQ has very big return in 2017, %199.4. 

However it looks like market has a significant turnover in 2018. Unless you invested into ENPH and RUN you could not be happy with the results. RUN had a significant move from 2017 to 2018. Their return was only %5.5 in 2017 however in 2018 the return was %84. The only stock that lose money for 2 years was TERP.

### The Code

On the Ecel VBA under Module1 I wrote the sub routine for All Stocks Analysis. Please see the picture below 

![allstock](https://user-images.githubusercontent.com/104239978/183552584-9f47c78f-a4f6-4305-9289-b4fe167ae719.png)

And I had to write a separete code to format output data table on the worksheet. Please see the picture below:

![allstock formatting](https://user-images.githubusercontent.com/104239978/183552629-b3b93933-e790-462f-b4ce-82dd74725830.png)

With this way, running the macro wasnt smooth enough. It was taking 1.070313 seconds for 2017 and 1.0625 seconfs for 2018 and also need to add on more for formatting the table.

To be able to run both macros together and save some time I refactored the previous code under Module3. Please see the picture below:

![allstock refactored](https://user-images.githubusercontent.com/104239978/183552666-a72ff350-9d7b-4e31-be9b-f73e6e3ea07d.png)

I created a timer, basic headers and initialized the list of the stocks that we wanted to analys. Than I created a ticker index and 3 output array and set the tickervolume to 0.

![allstock refactored1](https://user-images.githubusercontent.com/104239978/183552707-98519c89-b6a4-4c27-a8c5-0d005773e013.png)

Than I decided to work with Loop code and add the formatting code. This way code can scan all the rows data in the worksheet and put on the output table with formatted version.

After I run the refactored code, I was able to see the differences in timing. Now running the macro takes only 0.9375 seconds for 2017 and 0.90625 seconds for 2018. That is very great improvement.

## Summary

Refactoring helps code to become easy to run and understand. We can say by refactoring you preapere a ‘clean short cut’. Prepares better code to read easily and this will help your code to run faster. This is also helpful for debugging.

However if we are working with a very large code set, refactoring may become a big disadvantage. During the refactoring if we make any mistakes then we would spend more time on debugging the code.

With Excel VBA we can automate any functionality, it usually saves us time. However VBA scripting can be done only on excel so you need to learn the language.

