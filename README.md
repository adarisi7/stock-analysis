# stock-analysis
Anaylsis trends on stocks and tickers through VBA

## Overview 
Steve just graduated college and him and his parents want to invest in companies that involve green energy. We will help Steve and his family analysis the data with stocks using the VBA add-in software in excel. We analyze the total daily volume, starting and closing prices for every stock that their family is interested in. We compare the stocks for the years 2017 and 2018.

### Results
In 2017, we found that the stocks with the higest return was DQ, with the return of 199.4%. DQ also had the lowest Total Daily Volume, with it being 35,796,200. The lowest and only negative return in 2017 was by TERP. The higest total daily volume was by SEDG, with it being 782,187,000. The code for the original script ran for 0.6525 seconds for 2017. In the refactored code, the code ran for 0.1484375 seconds, which is clearly a lot faster than the original code. 

![2017!](https://github.com/adarisi7/stock-analysis/blob/289490b2926ea492eaa816ce219a8390f1e5ff7e/VBA_Challenge_2017.png)

In 2018, there were far more negative returns that positive returns. The only positive return was ENPH and RUN with returns of 84.0% and 81.9%. The stocks with the highest negative return was DQ, with -62.6%. The stock with the highest total daily volume was ENPH, and the lowest Total Daily Volume was the AY stock. The code in the original script ran for 0.625 seconds. In the refactored code, it ran for 0.140625 seconds. Again, it is clearly faster than the original script. 

![2018!](https://github.com/adarisi7/stock-analysis/blob/289490b2926ea492eaa816ce219a8390f1e5ff7e/VBA_Challenge_2018.png)

The advantage of the VBA refactored script is obviously that the running time is faster. Also the formatting is done automatically within the script. With the original script, it is not done so you have to do it in the formatting script. The advantage of the original script is that it is more organized to read and understand the code in case you went wrong, whereas in the refactored script it is harder because the code is more condensed. 


