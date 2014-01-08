excelsior
=========

Small convenience wrapper for PHPOffice/PHPExcel 

As working with PHPExcel can be pretty laborious at some point, I decided to create a small wrapper to make certain tasks more easy.
This library is actually just a small tool for personal use but maybe someone will find it usefull. Its in an early state and I will add more features as I need to.

### Classes
All classes actually represent wrapper around their PHPExcel equivalents whereas the Excelsior class represents the PHPExcel Class itself.

**Excelsior**

```php
$excel = new \Excelsior\Excelsior(); //will create an empty workbook
$excel = new \Excelsior\Excelsior('path/to/file'); //will load file instead

$sheet = $excel->getSheet(); //returns first sheet in workbook
```

**Workbook (Sheet)**  
Cells can either be fetched by coordinate or by column and row number.

```php
$cell = $sheet->getCell('C3');
//or
$cell = $sheet->getCell(2,3); //will return the same cell

```

**Cells**  
Excelsior is currently very cell-centric to simplify their handling.

Cell-Traversing
```php
$cell1 = $sheet->getCell('A1');
$cell2 = $cell->right(); //returns cell with coordinate A2
$cell3 = $cell->right(2); //returns cell with coordinate A3

//up, down, left methods are also available
```


Merging / Appending
```php
$cell1 = $sheet->getCell('A1');
$cell1 = $sheet->getCell('C3');

$cell1->merge($cell2); //merge the two cells to create a range of A1:C3

$cell3 = $sheet->getCell('H5');
$cell1->append($cell3); //the previously created range will be expanded to A1:H5

//Note that if any involved cell is part of a range, this range will be unmerged 
//prior to following merge or append operations
```

Styles
```php
$fill = array(
    'type' => \PHPExcel_Style_Fill::FILL_SOLID,
    'color' => array('rgb' => 'e6e6ff')
);

$cell = $sheet->getCell('A1');
$cell->setFill($fill); //applies given style to cell

$cell->merge($cell->down(4))->setFill($fill); //also works with ranges

//also available
$cell->setFont($configArray);
$cell->setBorders($configArray);
//see PHPExcel Documentation/Code for array-structure and available config options
```

TODO
----
*  add Tests
*  add more Examples
*  include some more features
