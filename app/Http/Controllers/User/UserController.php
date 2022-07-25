<?php

namespace App\Http\Controllers\User;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Exports\UsersExport;
use Maatwebsite\Excel\Facades\Excel;
use Illuminate\Support\Facades\File;

class UserController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function index()
    {
        //
    }

    /**
     * Show the form for creating a new resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function create()
    {
        //
    }

    /**
     * Store a newly created resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return \Illuminate\Http\Response
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function show($id)
    {
        //
    }

    /**
     * Show the form for editing the specified resource.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function edit($id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     *
     * @param  \Illuminate\Http\Request  $request
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function update(Request $request, $id)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     *
     * @param  int  $id
     * @return \Illuminate\Http\Response
     */
    public function destroy($id)
    {
        //
    }

    public function export()
    {
        $saveExcelPath = storage_path('excels');
        $completedPath = $saveExcelPath. DIRECTORY_SEPARATOR .'completed';
        // $this->createDirectory($completedPath);
        $excelFiles = scandir($saveExcelPath);

        foreach ($excelFiles as $key => $excelFile) {
            $extension = pathinfo($excelFile, PATHINFO_EXTENSION);
            $filename = pathinfo($excelFile, PATHINFO_FILENAME);

            if (!in_array($extension, ['xlsx', 'xls', 'csv'])) {
                continue;
            }

            // echo "processing $excelFile\n";

            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load(
                $saveExcelPath. DIRECTORY_SEPARATOR .$excelFile
            );

            $worksheets = $spreadsheet->getActiveSheet();
            $headers = $records = [];

            $i=0;
            foreach ($worksheets->toArray() as $num => $sheet) {
                foreach ($sheet as $row => $value) {
                    if ($num == 0) {
                        if (in_array($value, ['name', 'username'])) {
                            $value = 'name';
                        }

                        $headers[$row] = $value;
                        continue;
                    }
                    
                    $records[$i][$headers[$row]] = $value;
                }

                $i++;
            }

            return Excel::download(new UsersExport($records), 'aloha.xlsx');
        }
    }

    // public function createDirectory($path)
    // {
    //     if (!File::isDirectory($path)) {
    //         File::makeDirectory($path, 0755, true, true);
    //     }
    // }
}
