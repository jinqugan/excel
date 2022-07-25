<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Queue;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Illuminate\Support\Facades\File;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class BuildExcelFile extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'excel:dir';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Create an empty folder for excels if does not exist';

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $saveExcelPath = storage_path('excels');
        $completedPath = $saveExcelPath. DIRECTORY_SEPARATOR .'completed';
        $incompleteFile = $saveExcelPath. DIRECTORY_SEPARATOR .'incomplete';
        $created = $this->createDirectory($completedPath);
        if ($created) {
            echo "completed folder is created successfully at $completedPath\n";
        }

        $created = $this->createDirectory($incompleteFile);
        if ($created) {
            echo "incompleted folder is created successfully at $completedPath\n";
        }
    }

    private function createDirectory($path)
    {
        $created = false;

        if (!File::isDirectory($path)) {
            $created = true;
            File::makeDirectory($path, 0755, true, true);
        }

        return $created;
    }
}
