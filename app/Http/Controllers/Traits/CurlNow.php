<?php

namespace App\Http\Controllers\Traits;

use Illuminate\Http\Request;

trait CurlNow
{
    public function getcurl($input)
    {
        $curl = curl_init();
        $appKey = 'APP_KEY: '.$input['apiKey'];
        $user = 'LOGGED_USER: '.$input['user'];
        if (empty($input['flash'])) {
            $flash = 'FLASH: '.' ';
        }else {
            $flash = 'FLASH: '.$input['flash'];
        }

        if (empty($input['flash2'])) {
            $flash2 = 'FLASH2: '.' ';
        }else {
            $flash2 = 'FLASH2: '.$input['flash2'];
        }

        if (empty($input['flash3'])) {
            $flash3 = 'FLASH3: '.' ';
        }else {
            $flash3 = 'FLASH3: '.$input['flash3'];
        }
        curl_setopt_array($curl, array(
            CURLOPT_URL => config('services.url_wfp').$input['link'],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_TIMEOUT => 30000,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_HTTPHEADER => array(
                // Set Here Your Requesred Headers
                'Content-Type: application/json',
                $appKey,$user,$flash,$flash2,$flash3
            ),
        ));
        // return $response = curl_exec($curl);
        $response = curl_exec($curl);
        if(curl_error($curl)){
            abort(403, curl_error($curl));
        }
        curl_close($curl);
        $decoded = json_decode($response, true);
        return $decoded;
    }

    public function HRGet($input)
    {
        $curl = curl_init();
        $appKey = 'APP_KEY: '.$input['apiKey'];
        $user = 'LOGGED_USER: '.$input['user'];
        curl_setopt_array($curl, array(
            CURLOPT_URL => config('services.url_hr').$input['link'],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_CONNECTTIMEOUT => 3000,
            CURLOPT_TIMEOUT => 1,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_HTTPHEADER => array(
                // Set Here Your Requesred Headers
                'Content-Type: application/json',
                $appKey,$user
            ),
        ));
        $response = curl_exec($curl);

        if(curl_error($curl)){
            if(curl_errno($curl) == 28){
                 return $response ='time_out';
            }
            // abort(403, curl_errno($curl));
        }
        curl_close($curl);
        $decoded = json_decode($response, true);
        return $decoded;
    }

    public function getdebug($input)
    {
        $curl = curl_init();
        $appKey = 'APP_KEY: '.$input['apiKey'];
        $user = 'LOGGED_USER: '.$input['user'];
        if (empty($input['flash'])) {
            $flash = 'FLASH: '.' ';
        }else {
            $flash = 'FLASH: '.$input['flash'];
        }

        if (empty($input['flash2'])) {
            $flash2 = 'FLASH2: '.' ';
        }else {
            $flash2 = 'FLASH2: '.$input['flash2'];
        }

        if (empty($input['flash3'])) {
            $flash3 = 'FLASH3: '.' ';
        }else {
            $flash3 = 'FLASH3: '.$input['flash3'];
        }
        curl_setopt_array($curl, array(
            CURLOPT_URL => config('services.url_wfp').$input['link'],
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_TIMEOUT => 30000,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_HTTPHEADER => array(
                // Set Here Your Requesred Headers
                'Content-Type: application/json',
                $appKey,$user,$flash,$flash2,$flash3
            ),
        ));
        return $response = curl_exec($curl);
        $response = curl_exec($curl);
        if(curl_error($curl)){
            abort(403, curl_error($curl));
        }
        curl_close($curl);
        $decoded = json_decode($response, true);
        return $decoded;
    }
}
