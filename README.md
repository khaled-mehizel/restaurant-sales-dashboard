# restaurant-sales-dashboard
  Cleaning of a large dataset and building an interactive dashboard

# Overview 
This page will detail the process of cleaning a dataset and building a dashboard along with some VBA here and there as a part of the amazing [Kyle Pew](https://www.linkedin.com/in/kylepew/)'s [Microsoft Excel Data Analysis and Dashboard Reporting](https://www.udemy.com/course/microsoft-excel-data-analysis-and-dashboard-reporting/) course on Udemy.  
The workbook contains multiple sheets that we'll use to create sections of our dashboard. 

Customer Info sheet contains information about the restaurants: names, owners, contact info, location etc.
Order Info sheet contains information about the orders from said restaurants, when they were made, who shipped them and where etc.

# Resources and Tools Used 
- Microsoft Excel 
- Dataset provided by Kyle.

# Cleaning the dataset
- We verify the data types and dates in all the sheets.
- We format the data as tables and name them so that our formulas wouldn't get ruined if we add more records.
### Customer Info sheet
- We'll use text functions like **PROPER()** for names and **UPPER()** for ID's to keep the names and IDs consistent. We do this in a new column.
- We copy the values from the new column and paste them in the place of the old ones, so that there's the actual values instead of formula results.

### Order Info sheet
- We want to replace the numbers in the "ShipVia" column with the names of the shipping companies, we can **Find & Replace** or **CHOOSE()** which I prefer! I didn't know about it in my previous Excel project, it reminds me of a Python dictionary.
  - We feed the index argument with the numbers in the column, and the values are the shipping company names.
  - As usual, we copy the new values into the old column.

- We extract the month from the "OrderDate" column, so we can use it in our dashboard later, for aggregating order count per month for example.  This wouldn't be necessary in Power BI or Tableau!
  - Used **TEXT()** function to this end.
- Getting rid of the redundant columns like Shipping Address, Postal Code etc.

# Building the Dashboard
Added boxes for the customer name, their contact information, location, order count, and their order history.
## Customer Information
### Customer Name
Using **Data Validation**, we add a drop down menu containing all the customer names from the Customer Info worksheet.
### Contact Information
The contact info box contains the name of the owner, their phone number, and fax number. A simple **XLOOKUP()** will go very far here. For demonstration purposes, we'll use a **VLOOKUP()** and an **INDEX() MATCH()** combo as well.

I restricted the use of VLOOKUP() and INDEX MATCH to columns that don't have a missing value, since XLOOKUP has an argument for when if the value is missing, saving us the use of an **IF()** (Microsoft did great with XLOOKUP).

### Location Information
This section contains the restaurant's location information, like address, city, region (if applicable), and postal code.

Same as last time, I use **VLOOKUP()** and **INDEX() MATCH()** for columns that don't have missing values, and **XLOOKUP()** for the rest.

## Order History List

We copy the cleaned Order Information Table under the Order History List header.
We use **Advanced Filter** to make Excel filter the Order History according to the selected Customer Name. But we'll need to rerun the filter every time, this is where VBA will come in:
  - Ensure that the criteria and list arguments in the filter are accurate, then I simply record myself applying the filter in a macro.
  - We go into the VBA window and transport the code from its module into the appropriate worksheet, and paste into the window, and it'll trigger every time the user changes something in the worksheet.
  - To change it to only triggering when the customer name is changed, we add an **IF** statement that checks where the target cell is B3 (which is where the customer name is located).
  - The Order History List is now fully automated!



