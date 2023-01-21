# Excel to website auto input
A tool to automate the process of filling data into a website

## Steps:
1. Open **data_source.xlsx**
2. Fill in the target URL on cell B9
3. FIll in the actions sequentially starting from row 12 (notes: don't fill black cells)
4. Fill in the target Xpaths [Google it](https://letmegooglethat.com/?q=how+to+get+element+xpath)

## Notes:
1. Source column name is case sensitive
2. Actions will be repeated for each row on **data** sheet (if excel_data and/or excel_click are used)
3. Row 1 in **data** sheet is used for header/column name
4. Chromedriver version should match the existing Chrome browser on your machine [Download Chromedriver](https://chromedriver.chromium.org/)
