<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use Illuminate\Support\Arr;

ini_set('memory_limit', '2048M');

class ExcelMultiSheet extends Command
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

        $excelFiles = array_flip($excelFiles);
        unset($excelFiles['.']);
        unset($excelFiles['..']);
        unset($excelFiles['.DS_Store']);
        $excelFiles = array_flip($excelFiles);
        $totalMb = 0;

        try {
            if (empty($excelFiles)) {
                echo "no excel folder is found in $incompleteFile\n";
                return true;
            }

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
                $records = [];
                $subkey=0;
                $savedFiles = $completedPath. DIRECTORY_SEPARATOR .$excelFile.'.xlsx';

                if (!file_exists($savedFiles)) {
                    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter(
                        new Spreadsheet(),
                        'Xlsx'
                    );

                    $writer->save($savedFiles);
                }

                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                    $savedFiles
                );

                $worksheets = $spreadsheet->getActiveSheet();
                $rowBegin = $worksheets->getHighestDataRow();
                $totalFiles = count($subfiles);
                $currentFiles = 0;

                foreach ($subfiles as $subPos => $subfile) {
                    $extension = pathinfo($subfile, PATHINFO_EXTENSION);
                    $filename = pathinfo($subfile, PATHINFO_FILENAME);

                    if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                        $totalFiles--;
                        continue;
                    }

                    $spreadsheetsub = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                        $subpath. DIRECTORY_SEPARATOR .$subfile
                    );

                    $sheetCount = $spreadsheetsub->getSheetCount();
                    $currentFiles++;
                    for ($i = 0; $i < $sheetCount; $i++) {
                        $sheet = $spreadsheetsub->getSheet($i);
                        $sheetData = $sheet->toArray(null, true, true, true);
                        $sheetName = $sheet->getTitle();

                        // $worksheetsubs = $spreadsheetsub->getActiveSheet();
                        $headers = [];
                        $seq=null;
                        foreach ($sheetData as $num => $sheet) {
                            $check = array_filter($sheet);
                            if (empty($check)) {
                                continue;
                            }

                            $check = array_map('strtolower', $check);
                            $check = array_map('trim', $check);

                            if (is_null($seq) && in_array('name', $check)) {
                                $seq=0;
                            } elseif (is_null($seq)) {
                                continue;
                            }

                            foreach ($sheet as $row => $value) {
                                $value = strtolower(trim($value));

                                if ($seq == 0) {
                                    $value = $this->beautyHeader($value);

                                    if (!empty($value)) {
                                        $headers[$row] = $value;
                                    }

                                    continue;
                                }

                                if (!empty($headers[$row])) {
                                    if ($headers[$row] == 'name' && strlen($value) > 0) {

                                        while (strlen($value) > 0 && !ctype_alpha($value[0])) {
                                            $value = ltrim($value, $value[0]);
                                        }
                                    }

                                    $records[$subkey][$headers[$row]] = $value;
                                }

                            }

                            $seq++;
                            $subkey++;
                        }

                        unset($sheetData);
                    }

                    echo "=== ($currentFiles\\$totalFiles) === \n";
                    echo "filename: $subfile\n";

                    if (!empty($records)) {
                        $mb = mb_strlen(serialize($records), '8bit');
                        $totalMb += $mb;
                        $currentRecords = count($records);

                        echo "size(mb): $mb\n";
                        echo "total size(mb): $totalMb\n";
                        echo "current + previous = total (record) : $currentRecords + $rowBegin = ".( $rowBegin + $currentRecords)."\n";

                        $row = $rowBegin <= 0 ? 1 : $rowBegin;
                        $alpha = 'A';

                        if ($row == 1) {
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
                        }

                        foreach (($records) as $key => $record) {
                            if (empty($record['name'])) {
                                continue;
                            }

                            $row++;
                            $alpha = 'A';

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

                        $writer->save($savedFiles);
                        echo "=== end of ($currentFiles\\$totalFiles) === \n";
                        echo "done exported (folder:$excelFile) to $savedFiles\n\n";

                        $rowBegin = $row;
                    }

                    $records = [];
                }
            }

        } catch(\Exception $ex) {
            echo "sheetName: $sheetName\n";
            echo "something went wrong at line: ".$ex->getLine()."\n";
            echo $ex->getMessage();
            echo "\n";
            return false;
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
