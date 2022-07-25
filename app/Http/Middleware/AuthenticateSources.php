<?php

namespace App\Http\Middleware;

use Closure;
use App\Traits\ResponseTrait;

class AuthenticateSources
{
    use ResponseTrait;

    /**
     * Get the path the user should be redirected to when they are not authenticated.
     *
     * @param  \Illuminate\Http\Request  $request
     * @return string
     */
    public function handle($request, Closure $next)
    {
        $sources = config('constant.source_authentication');

        // if (!in_array($request->header(HEADER_SOURCE), config('constant.source_authentication'))) {
        if (!empty($sources[HEADER_SOURCE])) {
            $result = $this->requestErrors();
            $result['message'] = trans('general.header_source_invalid');
            $result['errors'] = ['source' => 'Please refer to technical team on the source value.'];

            return response()->json($result, $this->notFoundStatus);
        }

        return $next($request);
    }
}

