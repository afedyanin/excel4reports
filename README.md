# Simple Report Engine

Simple template based report engine

## Excel Reports

Generate Excel report file from DataSet object using template excel file.

### Books report sample

Template:

![BooksTemplate](/images/books_template.png)

Result:

![BooksReport](/images/books_report.png)

### Orders report sample

![OrdersTemplate](/images/orders_template.png)

Result:

![OrdersReport](/images/orders_report.png)

This project based on this article [EPPlus Excel Template Report Engine](https://www.codeproject.com/Articles/1252390/EPPlus-Excel-Template-Report-Engine)

But EPPlus replaced with [NPOI](https://github.com/tonyqus/npoi)

More info about NPOI

- [Getting Started with NPOI](https://github.com/nissl-lab/npoi/wiki/Getting-Started-with-NPOI)

- [Generate XLS from template](https://github.com/nissl-lab/npoi/blob/master/examples/hssf/GenerateXlsFromXlsTemplate/Program.cs)

- [Busy Developers' Guide to HSSF and XSSF Features](http://poi.apache.org/components/spreadsheet/quick-guide.html)

## Console Reports

Generate simple Text Table from DataSet object 

### Books

![BooksText](/images/books_txt.png)

### Orders

![OrdersText](/images/orders_txt.png)

Report builder uses [ConsoleTable](https://github.com/khalidabuhakmeh/ConsoleTables)

## Sample Code

### Sample App

[Program.cs](https://github.com/afedyanin/excel4reports/blob/main/src/SampleApp/Program.cs)

### Books

[BooksReport.cs](https://github.com/afedyanin/excel4reports/blob/main/src/SampleApp/Books/BooksReport.cs)

### Orders

[OrdersReport.cs](https://github.com/afedyanin/excel4reports/blob/main/src/SampleApp/Orders/OrdersReport.cs)

## Sample Data Sources

- [Orders sample data](https://www.contextures.com/xlsampledata01.html)

- [Books sample data](https://gist.github.com/nanotaboada/6396437)



