# Excel2MultipleCSVs
Split an Excel file with multiple tabs into CSV files for each tab. For Example: 

Input: 
-  Workbook.xlsx
    - tab1 = myfavouritecats
    - tab2 = myfavouritedogs
    
Output: 
- myfavouritecats.csv
- myfavouritedogs.csv


## Requirements
- Composer
- PHP 7.2 or greater

## Setup 
```bash
git clone https://github.com/LunarDevelopment/Excel2MultipleCSVs.git Excel2MultipleCSVs
cd Excel2MultipleCSVs 
composer install 
```

## Usage 
```bash
php Excel2CSVs.php --input="./AnExcelFile.xlsx" --output="./YourDirectoryToOutput"
```

