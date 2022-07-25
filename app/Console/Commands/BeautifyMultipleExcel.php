<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class BeautifyMultipleExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'format:multi';

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

        $spreadsheet = new Spreadsheet();
        $worksheets = $spreadsheet->getActiveSheet();

        foreach ($excelFiles as $key => $excelFile) {
            $extension = pathinfo($excelFile, PATHINFO_EXTENSION);
            $filename = pathinfo($excelFile, PATHINFO_FILENAME);

            if (in_array($excelFile, ['.', '..', '.DS_Store'])) {
                unset($excelFiles[$key]);
                continue;
            }

            $subpath = $incompleteFile. DIRECTORY_SEPARATOR .$excelFile;
            if (!is_dir($subpath)) {
                continue;
            }

            $subfiles = scandir($subpath);
            $headerScans = [];
            $records = [];
            $subkey=0;
            foreach ($subfiles as $key => $subfile) {
                $extension = pathinfo($subfile, PATHINFO_EXTENSION);
                $filename = pathinfo($subfile, PATHINFO_FILENAME);

                if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                    continue;
                }
                // echo "##############processing $subfile\n";
                // echo "subpathfile: ".$subpath. DIRECTORY_SEPARATOR .$subfile." \n";exit;

                $spreadsheetsub = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                    $subpath. DIRECTORY_SEPARATOR .$subfile
                );
                
                $worksheetsubs = $spreadsheetsub->getActiveSheet();
                $headers = [];
                $seq=0;
                foreach ($worksheetsubs->toArray() as $num => $sheet) {
                    
                    if (count(array_filter($sheet)) <= 3) {
                        continue;
                    }
                    
                    foreach ($sheet as $row => $value) {
                        $value = strtolower(trim($value));

                        if ($seq == 0) {
                            $value = $this->beautyHeader($value);

                            if (!empty($value)) {
                                $headers[$row] = $value;
                                $headerScans[$value] = $value;
                            }

                            continue;
                        }
                        
                        if (!empty($headers[$row])) {
                            $records[$subkey][$headers[$row]] = $value;
                        }

                    }

                    $seq++;
                    $subkey++;
                }
            }

            if (empty($records)) {
                echo "no excels record found\n";
                exit;
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

            
            foreach (($records) as $key => $record) {
                $row++;
                $alpha = 'A';

                if (empty($record['name'])) {
                    print_r($record);
                }

                $worksheets->getCell($alpha++ . $row)->setValue($record['name'] ?? null);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['nric'] ?? null), 's');
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['old_nric'] ?? null), 's');
                $worksheets->getCell($alpha++ . $row)->setValueExplicit($record['address1'] ?? null, DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit($record['address2'] ?? null, DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit($record['address3'] ?? null, DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValue($record['city'] ?? null);
                $worksheets->getCell($alpha++ . $row)->setValue($record['post'] ?? null);
                $worksheets->getCell($alpha++ . $row)->setValue($record['state'] ?? null);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['mobile1'] ?? null), DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['mobile2'] ?? null), DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['mobile3'] ?? null), DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['mobile4'] ?? null), DataType::TYPE_STRING);
                $worksheets->getCell($alpha++ . $row)->setValueExplicit(($record['mobile5'] ?? null), DataType::TYPE_STRING);
                $worksheets->getCell($alpha . $row)->setValue($record['category'] ?? null);
            }

            foreach (range('A', $alpha) as $key => $column) {
                $worksheets->getColumnDimension($column)->setAutoSize(true);
            }

            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter(
                $spreadsheet,
                'Xlsx'
            );

            $writer->save($completedPath. DIRECTORY_SEPARATOR .$excelFile.'.xlsx');
            echo "done exported $excelFile to $completedPath\n\n";
        }

        if (empty($excelFiles)) {
            echo "no excels found in $incompleteFile\n";
            exit;
        }
    }

    private function beautyHeader($value)
    {
        if (empty($value)) {
            return null;
        }

        $value = preg_replace('/\s+/', '', strtolower($value));
        $names = ['name', 'username'];
        $address1 = ['add1', 'address1'];
        $address2 = ['add2', 'address2'];
        $address3 = ['add3', 'address3'];
        $nric = ['icno', 'nric', 'ic'];
        $nricOld = ['oldicno', 'oldnric', 'oldic'];
        $mobile1 = ['mob1', 'mobile1', 'mobileno1', 'phone1', 'phoneno1'];
        $mobile2 = ['mob2', 'mobile2', 'mobileno2', 'phone2', 'phoneno2'];
        $mobile3 = ['mob3', 'mobile3', 'mobileno3', 'phone3', 'phoneno3'];
        $mobile4 = ['mob4', 'mobile4', 'mobileno4', 'phone4', 'phoneno4'];
        $mobile5 = ['mob5', 'mobile5', 'mobileno5', 'phone5', 'phoneno5'];

        if (in_array($value, $names)) {
            $value = 'name';
        }
        
        if (in_array(strtolower($value), $address1)) {
            $value = 'address1';
        }

        if (in_array(strtolower($value), $address2)) {
            $value = 'address2';
        }

        if (in_array(strtolower($value), $address3)) {
            $value = 'address3';
        }

        if (in_array(strtolower($value), $nric)) {
            $value = 'nric';
        }
        if (in_array(strtolower($value), $nricOld)) {
            $value = 'old_nric';
        }

        if (in_array(strtolower($value), $mobile1)) {
            $value = 'mobile1';
        }
        if (in_array(strtolower($value), $mobile2)) {
            $value = 'mobile2';
        }
        if (in_array(strtolower($value), $mobile3)) {
            $value = 'mobile3';
        }
        if (in_array(strtolower($value), $mobile4)) {
            $value = 'mobile4';
        }
        if (in_array(strtolower($value), $mobile5)) {
            $value = 'mobile5';
        }

        return $value;
    }

    private function createDirectory($path)
    {
        if (!File::isDirectory($path)) {
            File::makeDirectory($path, 0755, true, true);
        }
    }
}
