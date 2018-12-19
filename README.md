# VBA-of-Wall-Street

## Solutions:

### VBA Script:

- Wallstreet_Analysis - Hard.bas

### Screenshots:

- Wallstreet Analysis - Hard Screenshot - 2014.png
- Wallstreet Analysis - Hard Screenshot - 2015.png
- Wallstreet Analysis - Hard Screenshot - 2016.png

# Unit 2 | Assignment - The VBA of Wall Street

## Background

This project uses VBA scripting to analyze real stock market data.

### Files

- This [Test Data](Resources/alphabtical_testing.xlsx) is used in developing the scripts.

- Stock Data is used to generate the final report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

### Step one

- The script will loop through each year of stock data and grab the total amount of volume each stock had over the year.

- The code display the ticker symbol to coincide with the total volume.

- The result looks as follows.

![easy_solution](Images/easy_solution.png)

### Step Two

- The script that will loop through all the stocks and take the following info.

  - Yearly change from what the stock opened the year at to what the closing price was.

  - The percent change from the what it opened the year at to what it closed.

  - The total Volume of the stock

  - Ticker symbol

- The script also have conditional formatting that will highlight positive change in green and negative change in red.

- The result looks as follows.

![moderate_solution](Images/moderate_solution.png)

### Step Three

- The script is also able to locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".

- The result looks as follows.

![hard_solution](Images/hard_solution.png)

- The `alphabetical_testing.xlsx` is used while developing your code. This dataset is smaller which allows faster test.
