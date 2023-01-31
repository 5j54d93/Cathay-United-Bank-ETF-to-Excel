# Cathay United Bank ETF to Excel

![GitHub](https://img.shields.io/github/license/5j54d93/Cathay-United-Bank-ETF-to-Excel)
[![GitHub stars](https://img.shields.io/github/stars/5j54d93/Cathay-United-Bank-ETF-to-Excel)](https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/stargazers)
![GitHub repo size](https://img.shields.io/github/repo-size/5j54d93/Cathay-United-Bank-ETF-to-Excel)

Using [**Selenium**](https://www.selenium.dev) to extract ETF data from [**Cathay United Bank**](https://www.cathaybk.com.tw/cathaybk/personal/investment/etf/search/?hotissue=05) to Excel file with [**Safari**](https://www.apple.com/safari/).

<img src="https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/.github/assets/Cathay%20United%20Bank%20EFT.png" width='100%' height='100%'/>

## Flow Chart

<img src="https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/.github/assets/flow%20chart.png" width='100%' height='100%'/>

1. Open [EFT website from Cathay United Bank](https://www.cathaybk.com.tw/cathaybk/personal/investment/etf/search/?hotissue=05) with Safari and sleep 9 second
2. Find `table`、`thead`、`tbody`
3. Extract column title from `thead`
4. Extract rows from `tbody`
5. New or open a excel file named _Cathay-United-Bank-ETF.xlsx_
6. Set column title height、Set ETF name and price width
7. Set column title's font to bold
8. Write rows from `tbody` to Excel without tags that with `style='diaplay: none'`
9. If data is negative, set font color to red, else black
10. Save Excel file

## License：MIT

This package is [MIT licensed](https://github.com/5j54d93/Cathay-United-Bank-ETF-to-Excel/blob/main/LICENSE).
