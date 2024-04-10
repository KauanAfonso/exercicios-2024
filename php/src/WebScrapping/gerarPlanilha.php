
<?php

require_once '../../vendor/autoload.php';

use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\Style\BorderBuilder;
use Box\Spout\Writer\Common\Creator\Style\Color;

// Função para extrair os dados do HTML e escrever na planilha
// Função para extrair os dados do HTML e escrever na planilha
function extractAndWriteData($htmlFilePath, $outputFilePath) {
    $html = file_get_contents($htmlFilePath);

    $dom = new DOMDocument();
    @$dom->loadHTML($html); // O '@' é usado para suprimir erros de parsing

    // Localiza todos os elementos <a> com a classe "paper-card"
    $paperCards = $dom->getElementsByTagName('a');
    

    $writer = WriterEntityFactory::createXLSXWriter();
    $writer->openToFile($outputFilePath);

    

    // Cria um estilo para as células
    $style = (new StyleBuilder())
        ->setFontBold()
        ->setFontSize(12)
        ->setFontColor('FF12346') // Definindo a cor como azul (código RGB)
        ->setShouldWrapText(true)
        ->build();

    // Adiciona o cabeçalho da planilha com o estilo definido
    $headerRow = WriterEntityFactory::createRowFromArray(["Título", "Autores"], $style);
    $writer->addRow($headerRow);

    foreach ($paperCards as $paperCard) {
        // Extrai informações do título e autores
        $titleNode = $paperCard->getElementsByTagName('h4')->item(0);
        $title = $titleNode ? $titleNode->nodeValue : '';

        $authorsNode = $paperCard->getElementsByTagName('span');
        $authors = "";
        foreach ($authorsNode as $span) {
            $authors .= $span->nodeValue . ", ";
        }
        $authors = rtrim($authors, ", "); // Remove a última vírgula e o espaço

        // Adiciona as informações como uma nova linha na planilha com o estilo definido
        $row = WriterEntityFactory::createRowFromArray([$title, $authors], $style);
        $writer->addRow($row);
    }

    $writer->close();
}


// Caminho do arquivo HTML de origem
$htmlFilePath = './origin.html';
// Caminho para salvar o arquivo de saída (planilha)
$outputFilePath = './planilha.xlsx';

// Extrai dados e escreve na planilha
extractAndWriteData($htmlFilePath, $outputFilePath);


?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Download Planilha</title>
</head>
<body>
    <h1>Download Planilha</h1>
<?php


echo "Extração e escrita concluídas";?>

<a href="planilha.xlsx" download><button>Baixar Planilha</button></a>
</body>
</html>


