**customer orders.xlsm** has a spreadsheet containing customer orders ($), including one-time and repeat orders, and has a VBA module that allows the spreadsheet user to search for customers who, in total, spent _more_ than any given benchmark.

To begin, click the button on the spreadsheet to input a numerical amount as the benchmark. The VBA program sifts out customers whose total spend is greater than the benchmark, and then displays these customer accounts on a new worksheet called Report, showing each customer's ID and their respective total spend in table format.

Note: If a customer made repeat orders, the VBA program is designed to sum all the customer's orders before comparing it to the user input benchmark.


The VBA program can also:
- Perform user input validation.
- Sort the customer orders based on attributes such as date and ID.

To directly view VBA code, see **CustomerOrders.bas**.
