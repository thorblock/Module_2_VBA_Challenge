# Module 2 Challenge for Data Analytics
VBA focuses on data organization and summarization of stock data.

## TB Solutions
It contains the VBA separate exports and the images of my solution. The workbook is very *big*.
I ended up separating a lot of the pieces into separate modules. So to make it even more confusing, there is another macro whose sole purpose is to call the other macros.
A refactor might be in this project's future.

### VBA Bits
The module organization is not great, so here's a guide:

Module 2 makes up the bulk of it and creates the main summary table with each unique item, Yearly Change, Year Percent Change, and the Total Volume of each Ticker item. I used three If statements to track the total volume, the id of an opening price, and the id of a closing price.

Module 4 takes care of the formatting of the table created by Module 2. Red interior for cells less than or equal to zero and a green interior for everything else (if it's not negative or zero, it's positive). Additional If not needed, as the resulting formatting is binary anyway.

Module 5 is the summary of summary. Selecting for tickers with the most significant yearly % change, the lowest yearly % change, and the most significant volume. I used ranges first to identify the min and max values of the required metrics and then used those values as my comparison to find the ticker ids.

Module 3 calls the previous functions to run across all sheets in the book.

Module 1 is the eight-digit date conversion to mm/dd/yyyy. I'm unsure if this was an aspect of the assignment. I could not figure out As Date and DateSerial for the life of me. Module 1 is very much a "work the system" solution if the system was a box of fruit.
