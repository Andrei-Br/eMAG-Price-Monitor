# eMAG-Price-Monitor
Python script for checking the prices for products on https://www.emag.ro/ and writing the data in an Excel file. 

Libraries used: beautifulsoup4, lxml, openpyxl, datetime. 

### Step 1. Create the Excel file with the table header - Run TableHeader.py
![TableHeader](https://user-images.githubusercontent.com/48626600/64802058-82541200-d592-11e9-89a5-a5b39cd5d858.PNG)

Product - the name and the description of the product.

Store - either eMAG or partner store.

Date - the day and hour the data was scraped.

Link - the URL to the product.

### Step 2. Add the URLs of the products you want in OutputPrice.py
In the dictionary "links", the key is a short name of the product and the value is the URL 

![productURLs](https://user-images.githubusercontent.com/48626600/64803877-6a7e8d00-d596-11e9-85ff-cf654d802e80.PNG)


### Step 3. Run OutputPrice.py everytime you want to scrape the prices of the products in the Excel file. 

# Output Sampe
![OutputPrice](https://user-images.githubusercontent.com/48626600/64804297-39eb2300-d597-11e9-87e4-9dc8083be77e.PNG)
