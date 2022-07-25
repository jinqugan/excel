<?php

namespace App\Services;

use Illuminate\Support\Facades\Cache;
use App\Models\Status;

class UserService
{
    public function accountStatus(Array $params = [])
    {
        $statuses = Cache::remember('account_status', 5, function () {
            return Status::select('id', 'name')
            ->where(['type' => 'account'])
            ->get();
        });

        if (!empty($params['keyBy'])) {
            $statuses = $statuses->keyBy($params['keyBy']);
        }

        return $statuses->toArray();
    }
}
