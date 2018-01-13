<?php

use \Psr\Http\Message\ServerRequestInterface as Request;
use \Psr\Http\Message\ResponseInterface as Response;

use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Common\Type;
use Box\Spout\Writer\WriterFactory;

require '../vendor/autoload.php';
include_once '../include/Constants.php';
require_once '../Spout/Autoloader/autoload.php';

$app = new \Slim\App;

$app->get('/hello', function (Request $request, Response $response) {

    $response->getBody()->write("Hello, boss");
    return $response;
});

$app->get('/test', function (Request $request, Response $response) {
    $reader =ReaderFactory::create(Type::XLSX);
    $reader->open(FILE_PATH);

    $writer = WriterFactory::create(Type::XLSX); // for XLSX files

    $writer->openToFile(TEMP_FILE); // write data to a file or to a PHP stream
    // let's read the entire spreadsheet...
    foreach ($reader->getSheetIterator() as $sheetIndex => $sheet) {
        // Add sheets in the new file, as we read new sheets in the existing one
        if ($sheetIndex !== 1) {
            $writer->addNewSheetAndMakeItCurrent();
        }

        foreach ($sheet->getRowIterator() as $row) {
            // ... and copy each row into the new spreadsheet
            $writer->addRow($row);
        }
        $sheetName = $sheet->getName();
        $writer->getCurrentSheet()->setName($sheetName);
        if ($sheet->getName()===SHEET_NAME){
        // So let's add the new data:
            $writer->addRow(['2015-12-25', 'Christmas gift', 29, 'USD']);
            echo "Add ok\n";
        }
    }

    $reader->close();
    $writer->close();
    $response->getBody()->write("Write ok\n");
    unlink(FILE_PATH);
    rename(TEMP_FILE, FILE_PATH);

    return $response;
});

$app->get('/test2', function (Request $request, Response $response) {
    $writer = WriterFactory::create(Type::XLSX); // for XLSX files

    $writer->openToFile(TEMP_FILE); // write data to a file or to a PHP stream
    $writer->getCurrentSheet()->setName(SHEET_NAME);
    // let's read the entire spreadsheet...
     for($i=0; $i < 1000; $i++) {
        // Add sheets in the new file, as we read new sheets in the existing one
        $writer->addRow(['2015-12-25', 'Christmas gift', 29, 'USD', 'blabla', 'them vao cho nhieu', 164]);
    }

    $writer->close();
    $response->getBody()->write("Write ok\n");
    return $response;
});

$app->run();