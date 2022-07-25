<?php

namespace App\Exports;

use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\Exportable;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithColumnFormatting;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithTitle;
use Maatwebsite\Excel\Events\AfterSheet;
use Maatwebsite\Excel\Events\BeforeExport;
use Maatwebsite\Excel\Concerns\FromArray;
use App\Models\User;

class UsersExport implements
    FromCollection,
    WithMapping,
    WithHeadings
{
    use Exportable;

    protected $records;

    /**
     * Create a new controller instance.
     */
    public function __construct($records)
    {
        $this->records = $records;
    }

    /**
    * @return \Illuminate\Support\Collection
    */
    public function collection()
    {
        return collect($this->records);
    }

    /**
     * Write code on Method
     *
     * @return response()
     */
    public function headings() :array
    {
        return [
            'nric',
            'name',
            'date',
        ];
    }

    /**
     * Map for iterate result
     *
     * @var Invoice $payment
     */
    public function map($record): array
    {
        return [
            $record['nric'],
            $record['name'],
            $record['date'],
            $record['description'],
        ];
    }
}
