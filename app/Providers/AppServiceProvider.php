<?php

namespace App\Providers;

use Illuminate\Support\ServiceProvider;
use Illuminate\Support\Facades\Validator;

class AppServiceProvider extends ServiceProvider
{
    /**
     * Register any application services.
     *
     * @return void
     */
    public function register()
    {
        //
    }

    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot()
    {
        Validator::extend('checkspecialcharacter', function ($attribute, $value, $parameters, $validator)
        {
            $check = true;
            // $sp='"%*;<>?^`{|}~\\\'#=&';
            $sp = config('constant.special_character');

            if(preg_match("/[".$sp."]/",$value)) {
                $check=false;
            }

           return $check;
        });

        Validator::extend('alphabert', function ($attribute, $value, $parameters, $validator)
        {
            if (!ctype_alpha($value)) {
                return false;
            }

            return true;
        });
    }
}
