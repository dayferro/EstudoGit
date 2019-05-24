<?php

include('incs/lib.inc.php');
include('vendor/autoload.php');

use NumberToWords\NumberToWords;

// create the number to words "manager" class
$numberToWords = new NumberToWords();
$currencyTransformer = $numberToWords->getCurrencyTransformer('en');

if(empty($argv[1])) die("Modelo do documento não fornecido!\n");
$template = $argv[1];
if(!file_exists("template/$template")) die("Modelo do documento não encontrado!\n");

if(empty($argv[2])) die("CSV não fornecido!\n");
$csv = $argv[2];
if(!file_exists("csv/$csv")) die("csv não encontrado!\n");

if(empty($argv[3])) die("Padrao de saida nao fornecido!\n");
$outputPattern = $argv[3];

if (($handle = fopen("csv/$csv", "r")) === FALSE) die("Falha ao abrir o arquivo CSV!\n");

$colTitles = fgetcsv($handle, 1000, ";");

foreach($colTitles as $idx=>$title) $colTitles[$idx] = strtoupper($title);

if(!in_array("E-MAIL", $colTitles)) die("CSV nao possui campo email para disparo!\n");

$row = 1;
while (($data = fgetcsv($handle, 1000, ";")) !== FALSE) {
    $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor("template/$template");
    if(empty($tplVars)) $tplVars = $templateProcessor->getVariables();
    $emails = array();
    foreach($data as $idx=>$field) {
        if(in_array(strtoupper($colTitles[$idx]), $tplVars)){
            if(strtoupper($colTitles[$idx]) == 'VALOR'){
                $templateProcessor->setValue('VALOR', trim($data[$idx]) . " (" . valorPorExtenso(trim($data[$idx]), true, false) . ")");
                $valueIng = str_replace(array('.',','), '', $data[$idx]);
                $templateProcessor->setValue('VALOR-ING', trim($data[$idx]) . " (" . str_replace(array('dollars', 'dollar'), array('reais', 'real'), $currencyTransformer->toWords($valueIng, 'USD')) . ")");
            } else 
                $templateProcessor->setValue(strtoupper($colTitles[$idx]), trim($field));
        }
    }
    $emails[$row] = $data['E-MAIL'];
    $templateProcessor->saveAs("output/" . sprintf($outputPattern, $row) . ".docx");
    echo "\tGerado arquivo output/" . sprintf($outputPattern, $row) . ".docx\n";
    $row++;
 }

