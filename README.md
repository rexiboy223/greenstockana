This analysis was created with the purpose of refactoring the Microsoft Excel VBA code while obtaining better efficiency on the code's runtime. We obtain 12 different stock data over the years from 2017 and 2018. With this data, we want to determine which stocks would be advisable to invest in based on the stock's history. The daily stock data for each stock contains the following information:

Stock Opening price
Stock Closing price
Stock High Price
Stock Low Price
Stock Closing Price
Stock Adjusted Closing Price
Volume of the Stock
Arrays provide you with random access to elements, which allows you to access position elements more efficiently compared to data structures. For the refactored code I created 4 arrays which are called:

tickers(12) = This array intends to create a ticker symbol that will be used to identify each stock
tickerVolume(12) = This array intends to hold volume data
tickerStartingPrice(12) = This array intends to hold the starting price
tickerEndingPrice(12) = This array intends to hold the ending price
Our Data Results:
Looking solely at the 2017 results, it would be very misleading for a potential investor because 11 of the 12 stocks that we are analyzing are showing a positive return, and some stocks with close to a 200% increase. However, one year later in 2018 only 2 out of the 12 stocks were showing a positive return. Based on a two-year trend we can see that the following stocks return a positive return in both years. :

ENPH
RUN
2017 Stock Analysis:
