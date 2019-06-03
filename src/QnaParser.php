<?php

namespace Emerham;

use Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

final class QnaParser
{
    /**
     * The Main run function.
     *
     * @param int   $argc
     * @param array $argv
     *
     */
    public function runParse(int $argc, array $argv)
    {
        if ($argc < 3) {
            printf("Missing Input and Output Files\n");
            exit(1);
        }

        $inputFile = $argv[1];
        $outPutFile = $argv[2];

        $qnaFile = file_get_contents($inputFile);
        // Remove the BOM header that the QnA Bot adds in the bump.
        $bom = pack('H*', 'EFBBBF');
        $qnaFile = preg_replace("/^$bom/", '', $qnaFile);
//        $qnaFile = str_replace("\xEF\xBB\xBF", '', $qnaFile);
        if (self::isJson($qnaFile)) {
            $qnaFileParsed = json_decode($qnaFile, true);
        } else {
            printf("Input file is not valid JSON");
            exit(1);
        }
        // Create a new spreadsheet.
        $workBook = new Spreadsheet();
        // Parse all the questions and get an array back.
        $workSheets = self::parseQnaBotQuestions($qnaFileParsed);
        // Create the worksheets to use in the botium.
        self::createSheet($workSheets, $workBook);
        // Create a new Xlsx Object for writing.
        $writer = new Xlsx($workBook);
        // Save the workbook with the given file name.
        self::saveWorkbook($outPutFile, $writer);
        $workBook->disconnectWorksheets();
    }

    /**
     * Checks to see if a given string is valid JSON.
     *
     * @param string $string
     *
     * @return bool
     */
    private function isJson(string $string): bool
    {
        json_decode($string);

        return (json_last_error() == JSON_ERROR_NONE);
    }

    /**
     * Decode all the QnA Bot Data into an array.
     *
     * @param array $qnaData
     *
     * @return array
     */
    private function parseQnaBotQuestions(array $qnaData): array
    {
        $workSheets = [];
        foreach ($qnaData['qna'] as $qna) {
            if (empty($qna['t'])) {
                $workSheets["Default"][] = [
                    'q' => $qna['q'][0],
                    'a' => $qna['a'],
                ];
            } else {
                $workSheets[$qna['t']][] = [
                    'q' => $qna['q'][0],
                    'a' => $qna['a'],
                ];
            }
        }

        return $workSheets;
    }

    /**
     * Create the Spreadsheets grouped by the categories.
     *
     * @param array       $sheetData
     * @param Spreadsheet $workBook
     */
    private function createSheet(array $sheetData, Spreadsheet &$workBook)
    {
        $renamedFirstSheet = false;
        foreach ($sheetData as $group => $qna) {
            $row = 2;
            if (!$workBook->getSheetByName($group)) {
                if (!$renamedFirstSheet) {
                    $workBook->getActiveSheet()->setTitle($group);
                    $myTempSheet = $workBook->getActiveSheet();
                    $renamedFirstSheet = true;
                } else {
                    try {
                        $myTempSheet = $workBook->createSheet();
                        $myTempSheet->setTitle($group);
                    } catch (Exception $exception) {
                        printf(
                            "Failed to create a new sheet: %exception",
                            $exception
                        );
                        exit(1);
                    }
                }
                $myTempSheet->setCellValueByColumnAndRow(1, 1, 'Human');
                $myTempSheet->setCellValueByColumnAndRow(2, 1, 'Bot');
            }
            foreach ($qna as $section) {
                $myTempSheet = $workBook->getSheetByName($group);
                $myTempSheet->setCellValueByColumnAndRow(
                    1,
                    $row,
                    $section['q']
                );
                $row += 1;
                $myTempSheet->setCellValueByColumnAndRow(
                    2,
                    $row,
                    $section['a']
                );
                $row += 1;
            }
        }
    }

    /**
     * Save the workbook.
     *
     * @param string $outPutFile
     * @param Xlsx   $writer
     */
    private function saveWorkbook(string $outPutFile, Xlsx $writer)
    {
        try {
            $writer->save($outPutFile);
        } catch (Exception $exception) {
            echo sprintf(
                "Failed to save Workbook %s",
                $exception
            );
            exit(1);
        }
    }
}

