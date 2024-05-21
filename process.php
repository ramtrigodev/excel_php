<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file'])) {
    $fileTmpPath = $_FILES['file']['tmp_name'];
    $fileName = $_FILES['file']['name'];
    $fileSize = $_FILES['file']['size'];
    $fileType = $_FILES['file']['type'];
    $fileNameCmps = explode(".", $fileName);
    $fileExtension = strtolower(end($fileNameCmps));

    // Verifica se o arquivo é XLS
    $allowedfileExtensions = array('xls');
    if (in_array($fileExtension, $allowedfileExtensions)) {
        // Definir caminho para salvar o arquivo
        $uploadFileDir = './uploaded_files/';
        $dest_path = $uploadFileDir . $fileName;

        if(move_uploaded_file($fileTmpPath, $dest_path)) {
            echo "O arquivo foi carregado com sucesso.<br>";
            $inputFileName = $dest_path;

            try {
                $spreadsheet = IOFactory::load($inputFileName);
                $sheet = $spreadsheet->getActiveSheet();

                foreach ($sheet->getRowIterator() as $row) {
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);

                    foreach ($cellIterator as $cell) {
                        echo $cell->getValue() . "\t";
                    }
                    echo "<br>";
                }
            } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
                echo 'Erro ao carregar arquivo: ', $e->getMessage();
            }
        } else {
            echo 'Houve um problema ao carregar o arquivo.';
        }
    } else {
        echo 'Carregue um arquivo XLS válido.';
    }
} else {
    echo 'Nenhum arquivo foi carregado.';
}
?>
