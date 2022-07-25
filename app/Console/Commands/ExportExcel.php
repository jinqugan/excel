<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;
use App\Exports\UsersExport;
use Maatwebsite\Excel\Facades\Excel;

class ExportExcel extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'export:excel';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'export specific excel';

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
        $excelFiles = scandir($saveExcelPath);

        foreach ($excelFiles as $key => $excelFile) {
            $extension = pathinfo($excelFile, PATHINFO_EXTENSION);
            $filename = pathinfo($excelFile, PATHINFO_FILENAME);

            if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                continue;
            }

            echo "processing $excelFile\n";

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                $saveExcelPath. DIRECTORY_SEPARATOR .$excelFile
            );

            $worksheets = $spreadsheet->getActiveSheet();
            $headers = $records = [];

            foreach ($worksheets->toArray() as $num => $sheet) {
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

            return Excel::download(new UsersExport($records), $excelFile);
        }
    }

    private function createDirectory($path)
    {
        if (!File::isDirectory($path)) {
            File::makeDirectory($path, 0755, true, true);
        }
    }
}
