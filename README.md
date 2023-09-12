# What Does It Do?
The addin wants to help to create connected tables in Excel to Power BI.
The add-in is in its early development stages, but it currently offers:

- Connect to Power BI Desktop, Power BI Service, SSAS Tabular Models
- Create a query without DAX knowledge


![pbixl001](https://github.com/joschkos/pbixl/assets/50075326/471c05ea-1bf0-44fe-98f4-341605acba46)

# Getting Started
1. Download the right version of pbixl. Check which one fits to your Excel Installation (32Bit or 64Bit).

    [The Add-Ins](https://github.com/joschkos/pbixl/tree/main/Add-Ins)

2. The add-in can be installed by double-clicking the xll file in the zip file.
3. After opening a Power BI Desktop File click on the pbixl tab in excel.
4. Select the running Power BI Desktop instance and the query editor will be shown.

# Dependencies
This Add-In is made with [Excel-DNA](https://github.com/Excel-DNA)
The Add-In is referencing a data grid which is not part of this repository.

# Known Issues
The Add-In needs Excel 2016 or higher.
There is a bug if 2 instance of Power BI Desktop are open.
Please do not copy a table connected to Power BI into another Workbook. 

