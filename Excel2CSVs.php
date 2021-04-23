<?php
/**************************************************************************
 * Copyright (C) Lewis Dimmick, All Rights Reserved
 *
 * @file        Excel2CSVs.php
 * @author      Lewis Dimmick
 * @project     Excel2MultipleCSVs
 * @date        22/04/2021
 */


// Require composer's autoloader.
require_once 'vendor/autoload.php';

use Garden\Cli\Cli;
use \PhpOffice\PhpSpreadsheet\IOFactory;
use \PhpOffice\PhpSpreadsheet\Spreadsheet;

// Define the cli options.
$cli = new Cli();
$cli->description('Split an Excel File Into Multiple CSVs.')
    ->opt('input', 'Input File with Path.', true)
    ->opt('output', 'Output path for CSVs, must be a directory.', true);
// Parse and return cli args.
$args = $cli->parse($argv, true);
$input = $args->getOpt('input:i', './input.xlsx'); // get input file
$output = $args->getOpt('output:o', './output'); // get output dir

//$cli->description('Dump some information from your database.')
//    ->opt('host:h', 'Connect to host.', true)
//    ->opt('port:P', 'Port number to use.', false, 'integer')
//    ->opt('user:u', 'User for login if not current user.', true)
//    ->opt('password:p', 'Password to use when connecting to server.')
//    ->opt('database:d', 'The name of the database to dump.', true);

// parse workbook
$workbook = IOFactory::load($input);

foreach($workbook->getAllSheets() as $i => $sheet){
    $worksheet = $workbook->getSheet($i);
    $spreadsheet = new Spreadsheet();
    $spreadsheet->addSheet($worksheet, 0);
//    dump($spreadsheet);
//    continue;
    $filename = preg_replace('/[^a-zA-Z0-9_-]+/', '-', strtolower($sheet->getTitle())) . ".csv";
//    $writer = IOFactory::createWriter($spreadsheet->getParent(), "Csv"); // Just uses Education data set
    $writer = IOFactory::createWriter($spreadsheet, "Csv");
    $writer->save($output . "/". $filename);
    dump($sheet->getTitle() ) ;
}





//foreach($workbook->getAllSheets() as $sheet){
//    $spreadsheet = new Spreadsheet();
//    $spreadsheet->createSheet();
//    $spreadsheet->addSheet($sheet);
//    $spreadsheet->setActiveSheetIndex(0);
//    $filename = preg_replace('/[^a-zA-Z0-9_-]+/', '-', strtolower($sheet->getTitle())) . ".csv";
//
//    $writer = new Csv($spreadsheet);
//    $writer->save($output . "/". $filename);
//
////    $writer = IOFactory::createWriter($spreadsheet, "Csv");
////    $writer->save($output . "/". $filename);
//    dump($sheet->getTitle() ) ;
//    die();
//
//}
