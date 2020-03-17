<?php 
require "vendor/autoload.php";

$xls = new Arifrh\EasyExcel\EasyExcel('Example 1');

$xls->setColumnsWidth([
    'A' => 20, 'B' => 90
]);

$xls
    // merge cells with background color
    ->setColumnHeader('A1', 'B1', 'UTF-8 Support オートライブラリー', '90adf0')

    // set one cel
    ->setCellValue('A2', 'Easy Excel')

    // set multiple celss
    ->setCellValues([
        'B2' => 'Wrapper of PHPSpreadsheet',
        'B3' => 'Easy for simple use case'
    ])

    // set cell with background color
    ->setLabel('A4', 'Github Repository', 'e39054')

    // set cell with hyperlink
    ->setLinkValue('B4', 'https://github.com/arif-rh/easy-excel', 'Go to Github Repository')
    
    // clone sheet 
    ->cloneSheet('Example 1', 'Copy of Example 1')
    
    // or use name "Copy of Example 1"
    ->setActiveSheet(1)
    
    // insert 8 rows before row 1
    ->insertRows(8, 1)

    // add iamge to cell
    ->setImageValue('A1', './img/avatar.png')

    // set merge cells
    ->setMergeCellsValue('B1', 'B7', 'composer require arif-rh/easy-excel', ['HC', 'VC'])

    // download file, or add second parameter with 'pdf' to download as PDF
    ->forceDownload('Easy-Excel');