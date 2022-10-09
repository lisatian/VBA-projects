**customer orders.xlsm** has a spreasheet containing customer orders ($), including one-time and repeat orders, and has a VBA module that allows the spreasheet user to search for customers who, in total, spent _more_ than any given benchmark.

To begin, click the button on the spreasheet to input a numerical amount as the benchmark. The VBA program sifts out customers whose total spend is greater than the benchmark. These customer accounts get displayed on a new worksheet called Report, where a table displays each customer's ID and their respective total spend.

Note: If a customer has multiple order entries on the spreasheet, the VBA program is designed to sum the customer's orders before comparing it to the user input benchmark.


The VBA program can also:
- Perform user input validation.
- Sort the customer orders based on attributes such as date and ID.
