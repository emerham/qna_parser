<?php
include_once 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$workBook = new Spreadsheet();
$qnaFile = file_get_contents('round_1_data.json');
$qnaFileParsed = json_decode($qnaFile, TRUE);

$workSheets = [];
foreach ($qnaFileParsed['qna'] as $qna) {
  $workSheets[$qna['t']][] = [
    'q' => $qna['q'][0],
    'a' => $qna['a'],
  ];

}

foreach ($workSheets as $group => $qna) {
  $row = 2;
  if (!$workBook->getSheetByName($group)) {
    $myTempSheet = $workBook->createSheet();
    $myTempSheet->setTitle($group);
    $myTempSheet->setCellValueByColumnAndRow(1, 1, 'Human');
    $myTempSheet->setCellValueByColumnAndRow(2, 1, 'Bot');
  }
  foreach ($qna as $section) {
    $myTempSheet = $workBook->getSheetByName($group);
    $myTempSheet->setCellValueByColumnAndRow(1, $row, $section['q']);
    $row += 1;
    $myTempSheet->setCellValueByColumnAndRow(2, $row, $section['a']);
    $row += 1;
  }
}
$writer = new Xlsx($workBook);
$writer->save('myHello.xlsx');
$workBook->disconnectWorksheets();
