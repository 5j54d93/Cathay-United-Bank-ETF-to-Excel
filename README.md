# Cathay United Bank ETF to Excel

![GitHub](https://img.shields.io/github/license/5j54d93/Cathay-United-Bank-ETF-to-Excel)
[![GitHub stars](https://img.shields.io/github/stars/5j54d93/Cathay-United-Bank-ETF-to-Excel)](https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/stargazers)
![GitHub repo size](https://img.shields.io/github/repo-size/5j54d93/Cathay-United-Bank-ETF-to-Excel)

Using [**Selenium**](https://www.selenium.dev) to extract ETF data from [**Cathay United Bank**](https://www.cathaybk.com.tw/cathaybk/personal/investment/etf/search/?hotissue=05) to Excel file with [**Safari**](https://www.apple.com/safari/).

<img src="https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/.github/assets/Cathay%20United%20Bank%20EFT.png" width='100%' height='100%'/>

## Preparation

- A Mac, so you could use Safari
- Allow "remote automation" in Safari
> 1. Choose Safari > Preferences, and on the Advanced tab, select "Show Develop menu in menu bar."
> 2. Choose Develop > Allow Remote Automation.

## Flow Chart

<img src="https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/.github/assets/Flow%20Chart.png" width='100%' height='100%'/>

`try:`

> 1. Set up web driver：Safari
> 2. Open [EFT website on Cathay United Bank](https://www.cathaybk.com.tw/cathaybk/personal/investment/etf/search/?hotissue=05)
> 3. New or open a excel file named _Cathay-United-Bank-ETF.xlsx_
> 4. Initial Excel cell format
> 5. Wait for website to load completely：until price of the last row not equal to "-"
> 6. Extract `table`、`thead`、`tbody`
> 7. Write column title to Excel file, and set font to bold
> 8. Write table data to Excel file：add new line on EFT name cell and price cell
> 9. Save and close Excel file
> 10. Quit web driver

`except Exception:`

> 1. If an Excel file have been opened, close it
> 2. If a web driver have been created, quit it
> 3. Rerun _main.py_ by `os.execv(sys.executable, ['python3'] + sys.argv)`

## License：MIT

This package is [MIT licensed](https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/LICENSE).
