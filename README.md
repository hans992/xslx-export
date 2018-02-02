XLSX-export
===================

## Exporting data from database to xslx

### Given database structure


```
CREATE TABLE `delivery_stats` (
  `ID` int(10) UNSIGNED NOT NULL,
  `IDF` varchar(12) COLLATE utf8_unicode_ci NOT NULL,
  `method` varchar(20) COLLATE utf8_unicode_ci NOT NULL,
  `count` int(10) UNSIGNED NOT NULL,
  `date` date NOT NULL,
  `cartValue` float DEFAULT '0'
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;
```

### Given definition

It's necessary to generate form that contains date input fields and export to xslx button.
Xlsx file should contain:

* sheet 1:
    - 1. In first cell show data range for data selection from database
    - 2. Number of orders for every method in selected date range
    - 3. Total count of orders in selected data range
    - 4. Average price of orders from selected date range. Prices that are 0 are excluded from average calculation.

* sheet 2:
    - 1. Contains 3 columns
    - 2. Orders are grouped by method and month


Package used for generating xlsx file:
[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet#want-to-contribute) 