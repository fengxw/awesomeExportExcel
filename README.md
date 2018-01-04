## Awesome export excel

It's an easier way to export ant excel. you don't need to worry about the detail, the default style of excel is enough for you.

### Install

```console
composer require fengxw/awesome-export-excel
```

### Usage

```php
// define the arguments.
$header = 'It is an header';
$title = ['column1', 'column2'];
$data = [['value1', 'value2']];
$sheetName = 'sheet 1';

// export excel
$exportExcel = new ExportExcel();
$exportExcel->export($header, $title, $data, $sheetName);
```
