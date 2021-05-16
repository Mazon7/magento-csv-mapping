# magento-csv-mapping
Processing CSV files for Magento import

The app contains of few parts:
- GUI interface
- Logic of data processing
- Attributes file that works as a database for data processing
- Input file and output file

So the aim of the app is to automate the process of mapping different CSV files (goods suppliers) into one format for import into Magento system.
App is built using only functional approach that may seem a drawback, but for this specific case that worked.

For GUI the TkInter is used and for data processing openpyxl suited best. 
Since the app works with the very large csv data sets, the threading is also used for avoiding GUI glitches.