# Coffee-Data-Excel-Project
I developed a comprehensive project in MS Excel, creating pivot tables, charts and an interactive dashboard to analyze the data. This I achieved by carrying out data preprocessing and visualization.
# Introduction
## Objective
The purpose of this analysis was to study how different factors influence sales of coffee in the United States, Ireland and the United Kingdom, for a continuous period of time – from January 2019 to July 2022. From the analysis I was able to gain insights into the trend of sales and how these factors affect them. These factors include type of coffee (robusta, excelsa, and liberica), size of the coffe, roast type, and the presence or absence of a loyalty card.
## The data set collected contains 3 tables (orders, customers, and products) with the following variables:
Orders                         
* Customer ID	
* Product ID	Quantity	
* Customer Name	
* Email	
* Country	
* Coffee Type	
* Roast Type	Size	
* Unit Price	
* Sales

Customers
* Customer ID
* Customer Name
* Email	Phone Number
* Address Line 1	City
* Country	Postcode
* Loyalty Card

Products
* Product ID	
* Coffee Type	
* Roast Type	Size	
* Unit Price	Price per 100g	
* Profit
# Configuration
The data was clean and did not need much pre-processing.
# Data Features
I used different functions in excel to populate the empty fields in the order table so as to gain insight into the data. 
XLOOKUP
I used XLOOUP to find the exact matches for Customer Name, Email Address, Country and Loyalty Card in the order table to the customers table.
=XLOOKUP(C3,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)  - Customer Name
=IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0))) - Email Address
=XLOOKUP(orders!C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)  - Country
=XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0) - Loyalty Card

IF
I used IF to return the full names of the Coffee Types and Roast Types from their abbreviations.
=IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica", IF(I2="Lib", "Liberica", "")))) - this resulted to the new column Coffee Type Name.
=IF(J2="M", "Medium", IF(J2="L","Light", IF(J2="D","Dark"))) - this resulted to the new column Roast Type Name.

INDEX
I used INDEX in combination with MATCH to populate product fields (Coffee Type,	Roast Type,	Size, and	Unit Price) in the order table from the product table. I used Product ID  and name of the columns e.g. Coffee Type as lookup values from the orders table, and the fields Product ID,	Coffee Type,	Roast Type,	Size, and	Unit Price from the product table as my lookup arrays.
=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0)) - Coffee Type.
=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!J$1,products!$A$1:$G$1,0)) - Roast Type.
=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!K$1,products!$A$1:$G$1,0)) - Size.
=INDEX(products!$A$1:$G$49,MATCH(orders!$D3,products!$A$1:$A$49,0),MATCH(orders!L$1,products!$A$1:$G$1,0)) - Unit Price.

# Data Analysis
Different charts were used to visualize the data and draw meaning from them. Bar charts were used to visualize categorical variables of numerical data and line charts were used to show trends between different variables over a continuous period of time.

# Dashboard
![Dashboard](/coffee_sales_dashboard.png)




The data was pre-processed using 

