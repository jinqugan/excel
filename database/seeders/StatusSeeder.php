<?php

namespace Database\Seeders;

use Illuminate\Database\Console\Seeds\WithoutModelEvents;
use Illuminate\Database\Seeder;
use App\Models\Status;

class StatusSeeder extends Seeder
{
    /**
     * Run the database seeds.
     *
     * @return void
     */
    public function run()
    {
        $statuses = [
            /**
             * Account status
             */
            [
                'name' => 'pending',
                'lang_code' => 'status.account_pending',
                'type' => 'account',
                'access' => true,
                'description' => 'account valid and unverified. Account unverified able to login'
            ],
            [
                'name' => 'active',
                'lang_code' => 'status.account_active',
                'type' => 'account',
                'access' => true,
                'description' => 'account valid and verified. Account verified able to login'
            ],
            [
                'name' => 'inactive',
                'lang_code' => 'status.account_inactive',
                'type' => 'account',
                'access' => true,
                'description' => 'account valid but inactive for a period of time. Account inactive able to login'
            ],
            [
                'name' => 'disabled',
                'lang_code' => 'status.account_disabled',
                'type' => 'account',
                'access' => false,
                'description' => 'account invalid and disabled by admin. Account disabled unable to login'
            ],
            [
                'name' => 'suspended',
                'lang_code' => 'status.account_suspended',
                'type' => 'account',
                'access' => false,
                'description' => 'account invalid and suspended by admin/system. Account suspended unable to login'
            ],
            [
                'name' => 'deleted',
                'lang_code' => 'status.account_deleted',
                'type' => 'account',
                'access' => false,
                'description' => 'account invalid and not exist in our system. Account deleted unable to login'
            ],
            /**
             * Product status
             */
            [
                'name' => 'pending_publish',
                'lang_code' => 'status.product_pending_publish',
                'type' => 'product',
                'access' => false,
                'description' => 'product added to system, waiting for authorise admin to review and make approval'
            ],
            [
                'name' => 'published',
                'lang_code' => 'status.product_published',
                'type' => 'product',
                'access' => false,
                'description' => 'product added to system, reviewed and approved by admin'
            ],
            [
                'name' => 'draft',
                'lang_code' => 'status.product_draft',
                'type' => 'product',
                'access' => false,
                'description' => 'product details add partialy to system, incomplete and able to fill in next time',
            ],
            [
                'name' => 'out_of_stock',
                'lang_code' => 'status.product_outofstock',
                'type' => 'product',
                'access' => false,
                'description' => 'product details add partialy to system, incomplete and able to fill in next time',
            ],
            [
                'name' => 'deleted',
                'lang_code' => 'status.product_outofstock',
                'type' => 'product',
                'access' => false,
                'description' => 'product deleted but still exist in system, to display only in market',
            ]
        ];

        Status::insert($statuses);
    }
}
