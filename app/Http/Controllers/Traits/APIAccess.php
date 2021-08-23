<?php
namespace App\Http\Controllers\Traits;

trait APIAccess {
    function getAppkey($data)
    {
        $access = unserialize(base64_decode(session('access_tokens')));
        if (isset($access)) {
            switch (true) {
                case (isset($access[$data])):
                    return $access[$data];
                    break;
                default:
                    return 'NO_APP_KEY';
                    break;
            }
        }
    }
}