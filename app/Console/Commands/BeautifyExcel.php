<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;

class BeautifyExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'format:excel';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $saveExcelPath = storage_path('excels');
        $completedPath = $saveExcelPath. DIRECTORY_SEPARATOR .'completed';
        $this->createDirectory($completedPath);
        $incompleteFile = $saveExcelPath. DIRECTORY_SEPARATOR .'incomplete';
        $excelFiles = scandir($incompleteFile);

        foreach ($excelFiles as $key => $excelFile) {
            $extension = pathinfo($excelFile, PATHINFO_EXTENSION);
            $filename = pathinfo($excelFile, PATHINFO_FILENAME);

            if (in_array($excelFile, ['.', '..', '.DS_Store'])) {
                continue;
            }

            $subpath = $incompleteFile. DIRECTORY_SEPARATOR .$excelFile;
            if (!is_dir($subpath)) {
                continue;
            }

            $subfiles = scandir($subpath);
            $headerScans = [];
            foreach ($subfiles as $key => $subfile) {
                $extension = pathinfo($subfile, PATHINFO_EXTENSION);
                $filename = pathinfo($subfile, PATHINFO_FILENAME);

                if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                    continue;
                }
                echo "##############processing $subfile\n";
                // echo "subpathfile: ".$subpath. DIRECTORY_SEPARATOR .$subfile." \n";exit;

                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                    $subpath. DIRECTORY_SEPARATOR .$subfile
                );
                
                $worksheets = $spreadsheet->getActiveSheet();
                $headers = $records = [];
                $seq=0;
                foreach ($worksheets->toArray() as $num => $sheet) {
                   
                    if (count(array_filter($sheet)) <= 3) {
                        continue;
                    }

                    foreach ($sheet as $row => $value) {
                        $value = strtolower(trim($value));

                        // if ($value == 'addre') {
                        //     echo "subfile : end of add:\n";
                        //     print_r($subfile);
                        //     print_r($sheet); exit;
                        // }

                        if ($seq == 0) {
                            if (in_array($value, ['name', 'username'])) {
                                $value = 'name';
                            }

                            if (in_array(strtolower($value), ['add 1', 'address1'])) {
                                $value = 'add 1';
                            }

                            if (in_array(strtolower($value), ['add 2', 'address2'])) {
                                $value = 'add 2';
                            }

                            if (in_array(strtolower($value), ['add 3', 'address3'])) {
                                $value = 'add 3';
                            }

                            if (!empty($value)) {
                                $headers[$row] = $value;
                                $headerScans[$value] = $value;
                            }

                            continue;
                        }
                        
                        if (!empty($headers[$row])){
                            $records[$num][$headers[$row]] = $value;
                        }

                    }

                    $seq++;
                }

                $row = 1;
                $alpha = 'A';

                $worksheets->getCell($alpha++ . $row)->setValue('Name');
                $worksheets->getCell($alpha++ . $row)->setValue('IC No');
                $worksheets->getCell($alpha++ . $row)->setValue('Old IC');
                $worksheets->getCell($alpha++ . $row)->setValue('Add 1');
                $worksheets->getCell($alpha++ . $row)->setValue('Add 2');
                $worksheets->getCell($alpha++ . $row)->setValue('Add 3');
                $worksheets->getCell($alpha++ . $row)->setValue('City');
                $worksheets->getCell($alpha++ . $row)->setValue('Post');
                $worksheets->getCell($alpha++ . $row)->setValue('State');
                $worksheets->getCell($alpha++ . $row)->setValue('Mob 1');
                $worksheets->getCell($alpha++ . $row)->setValue('Mob 2');
                $worksheets->getCell($alpha++ . $row)->setValue('Mob 3');
                $worksheets->getCell($alpha++ . $row)->setValue('Mob 4');
                $worksheets->getCell($alpha++ . $row)->setValue('Mob 5');
                $worksheets->getCell($alpha++ . $row)->setValue('Category');

                foreach ($records as $key => $record) {
                    $row++;
                    $alpha = 'A';

                    $worksheets->getCell($alpha++ . $row)->setValue($record['name'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['ic no'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue("");
                    $worksheets->getCell($alpha++ . $row)->setValue($record['add 1'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['add 2'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['add 3'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['city'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['post'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['state'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['mob 1'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['mob 2'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['mob 3'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['mob 4'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['mob 5'] ?? null);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['category'] ?? null);
                }

                $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter(
                    $spreadsheet,
                    'Xlsx'
                );

                $writer->save($completedPath. DIRECTORY_SEPARATOR .$excelFile.'.xlsx');
                echo "done exported $excelFile to $completedPath\n\n";
                exit;
            }

            print_r($headerScans);exit;
        }
    }

    private function createDirectory($path)
    {
        if (!File::isDirectory($path)) {
            File::makeDirectory($path, 0755, true, true);
        }
    }
}
