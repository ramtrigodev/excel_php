<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Cell;

// Configurações do banco de dados
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "seu_banco_de_dados";

// Criar conexão com o banco de dados
$conn = new mysqli($servername, $username, $password, $dbname);

// Verificar conexão
if ($conn->connect_error) {
    die("Falha na conexão: " . $conn->connect_error);
}

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file'])) {
    $fileTmpPath = $_FILES['file']['tmp_name'];
    $fileName = $_FILES['file']['name'];
    $fileSize = $_FILES['file']['size'];
    $fileType = $_FILES['file']['type'];
    $fileNameCmps = explode(".", $fileName);
    $fileExtension = strtolower(end($fileNameCmps));

    // Verifica se o arquivo é XLS
    $allowedfileExtensions = array('xls', 'xlsx');
    if (in_array($fileExtension, $allowedfileExtensions)) {
        // Definir caminho para salvar o arquivo
        $uploadFileDir = './uploaded_files/';
        $dest_path = $uploadFileDir . $fileName;

        if (move_uploaded_file($fileTmpPath, $dest_path)) {
            echo "O arquivo foi carregado com sucesso.<br>";
            $inputFileName = $dest_path;

            try {
                $spreadsheet = IOFactory::load($inputFileName);
                $sheet = $spreadsheet->getActiveSheet();

                foreach ($sheet->getRowIterator() as $row) {
                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);

                    $coluna1 = '';
                    $coluna2 = '';
                    $coluna3 = '';
                    $coluna4 = '';

                    $i = 1;
                    foreach ($cellIterator as $cell) {
                        switch ($i) {
                            case 1:
                                $coluna1 = $cell->getValue();
                                break;
                            case 2:
                                $coluna2 = $cell->getValue();
                                break;
                            case 3:
                                $coluna3 = $cell->getValue();
                                break;
                            case 4:
                                $coluna4 = $cell->getValue();
                                break;
                            // Adicione mais casos conforme necessário
                        }
                        $i++;
                    }

                    $sql = "INSERT INTO dados_excel (coluna1, coluna2, coluna3, coluna4) VALUES (?, ?, ?, ?)";
                    $stmt = $conn->prepare($sql);
                    $stmt->bind_param("ssss", $coluna1, $coluna2, $coluna3, $coluna4);
                    $stmt->execute();
                }

                echo "Dados inseridos no banco de dados com sucesso.";
            } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
                echo 'Erro ao carregar arquivo: ', $e->getMessage();
            }
        } else {
            echo 'Houve um problema ao carregar o arquivo.';
        }
    } else {
        echo 'Carregue um arquivo XLS ou XLSX válido.';
    }
} else {
    echo 'Nenhum arquivo foi carregado.';
}

// Fechar conexão com o banco de dados
$conn->close();
?>

