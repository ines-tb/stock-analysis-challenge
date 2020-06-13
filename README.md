# Challenge : Stock Analysis

Analysis of stocks from different companies to find out their Total Volume and Yearly Return using Visual Basic for Application (VBA). 

### Data source
The dataset used in this challenge has the stock tickers data for twelve companies alphabetically ordered. That information is stored in a different sheet per year, named following the format 'yyyy'.

### Program
######  > Overview
The macro built for the analysis is located in _Module3_AllStocksChallenge_ in the xlsm file named _green_stocks_. It works as follows:
The user will be requested to input the year to be analyzed, then every non-null row, except for the header, in the corresponding sheet will be checked and the needed information stored accordingly for subsequent use in the population of the output sheet.
Additionally, some format will be applied to the data to make it easier to visually extract conclusions.
Note that for the analysis, it is been assumed that the number of tickers to be processed will be twelve.

######  > Variables
Apart from a few auxiliary variables, four different arrays of twelve elements have been created to store the data needed.
* **tickers** - To be populated with the ticker's symbols.
* **startingPrices** - To be populated with the closing prices of each company at the beginning of the year. 
* **endingPrices** - To be populated with the closing prices of each company at the end of the year.
* **totalVolumes** - To be populated with the total volume of the company for that year.
  
Note that the values of the elements in the last three arrays should correspond to the ticker in the exact same position in 'tickers' array.

###### > Body
After the variables creation, four loops and a block for static formatting will be executed.
* **Total volumes array initialization loop:**
  All elements in totalVolume are set to 0. As the totalVolume sums the previous volume amount for that ticker, the first time a new ticker is processed there is not previous amount and setting initially to 0 we avoid potential data type errors.  
* **Data check loop:** 
  From the second (to skip the header) until the last non-null row (previously calculated and stored in rowCount variable), each line in year-data sheet is checked.
  Along the loop, the positions where the data is stored in the arrays are controlled by the variable _tickerIndex_. Initially set to 0, it will be incrementing as long as a new ticker is found in the data, which means that next values will be stored in the following element in the four arrays.
  Once inside the loop,
  - When a new ticker is found, (_i.e. the value of the first column in current row differs from the one in previous row_), the name is stored in _tickers_ array in the position indicated by _tickerIndex_. Also, this means that the current row is the first occurrence of the ticker so the closing value (_sixth column in data sheet_) is stored in _startingPrices_ array.
  - Then, the volume for that day is added to the element in _totalVolume_ array corresponding to the current ticker (_totalVolume(tickerIndex)_). The first time, the volume found will be added to 0, as every element had been already initialized to that number.
  - Finally, it is checked whether this is the last occurrence of the current ticker (_i.e. the value of the first column in following row differs from current ticker's name_) store the closing value (_sixth column in data sheet_) in _endingPrices_ array. After that, _tickerIndex_ counter is incremented.


  
* **Output loop:**
  In previous loop all the needed data was stored in the created arrays.
  Here, after activating the sheet where the analysis output will be written, we loop as many times as the number of tickers. One row per ticker will be added with the name in the first column, its total volume in the second and the yearly revenue in the third. 

* **Static formatting:**
  A block with some static formatting sentences will change the font, alignment and number formats of certain elements of the output sheet. E.g. Including some color to the header, adding decimals, '$' and thousand separator to the volume or writing the '%' symbol at the end of the return values with certain decimals.

* **Conditional formatting:**
  A last loop will change the filling color of the third column (_Return_) depending on its value. Red for negative values, green for positive and none otherwise.

###### > Macro Output
In the output sheet the first cell of the worksheet will indicate that the analysis is for all stocks and the year for which the analysis has been run. 
A table below, will contain the summarized data for each ticker: the total volume and the yearly revenue percentage. The cells of these last values will be highlighted red or green for negative or positive values respectively.