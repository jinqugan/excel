<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;

class ScanExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'scan:excel';

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

            if (in_array($excelFile, ['.', '..'])) {
                continue;
            }

            $subpath = $incompleteFile. DIRECTORY_SEPARATOR .$excelFile;
            if (!is_dir($subpath)) {
                continue;
            }

            $subfiles = scandir($subpath);
            
            foreach ($subfiles as $key => $subfile) {
                echo "processing $subfile\n";
                $extension = pathinfo($subfile, PATHINFO_EXTENSION);
                $filename = pathinfo($subfile, PATHINFO_FILENAME);

                if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                    continue;
                }
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                    $subpath. DIRECTORY_SEPARATOR .$subfile
                );

                $worksheets = $spreadsheet->getActiveSheet();
                $headers = $records = [];

                $seq=0;
                foreach ($worksheets->toArray() as $num => $sheet) {
                    if (count($sheet) <= 3) {
                        print_r($sheet);exit;
                    }

                    foreach ($sheet as $row => $value) {
                        if ($num == 0) {
                            if (in_array($value, ['name', 'username'])) {
                                $value = 'name';
                            }

                            $headers[$row] = $value;
                            continue;
                        }
                        
                        $records[$num][$headers[$row]] = $value;
                    }
                }
                
                print_r($records);exit;
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

                    $worksheets->getCell($alpha++ . $row)->setValue($record['date']);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['nric']);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['name']);
                    $worksheets->getCell($alpha++ . $row)->setValue($record['description']);
                }

                $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter(
                    $spreadsheet,
                    'Xlsx'
                );

                $writer->save($completedPath. DIRECTORY_SEPARATOR .$excelFile);
                echo "done exported $excelFile to $completedPath\n\n";
            }

        }
    }

    private function createDirectory($path)
    {
        if (!File::isDirectory($path)) {
            File::makeDirectory($path, 0755, true, true);
        }
    }
}
