<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

$inputFileName = './arquivo.xls';

try {
    
    $spreadsheet = IOFactory::load($inputFileName);

    $sheet = $spreadsheet->getActiveSheet();

    foreach ($sheet->getRowIterator() as $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false); 

        foreach ($cellIterator as $cell) {
            echo $cell->getValue() . "\t";
        }
        echo "\n";
    }
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo 'Erro ao carregar arquivo: ', $e->getMessage();
}
?>
