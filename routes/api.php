<?php

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| is assigned the "api" middleware group. Enjoy building your API!
|
*/

Route::middleware('auth:api')->get('/user', function (Request $request) {
    return $request->user();
});

// Route::get('/exportWord', ['as'=>'createWord','uses'=>'WordTestController@createWordDocx']);
Route::group(['prefix' => 'export', 'as' => 'export'], function() {
    Route::get('/postqual/{year?}/{key?}' , 'WordController@postQualEvalReport')->name('.postqual');
    // Route::get('/pr/{year?}/{key?}' , 'ExcelController@purchaseRequest')->name('.pr');



    Route::get('/word' , 'WordController@index')->name('.word');
    Route::get('/sample' , 'WordController@sample')->name('.sample');

    Route::get('/id/generate' , 'IDController@generate')->name('.id.generate');

    Route::group(['prefix' => 'xls', 'as' => 'xls'], function() {
        Route::get('/pr' , 'ExcelController@purchaseRequest')->name('.pr');
        Route::get('/pr/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@purchaseRequest')->name('.pr');
        Route::get('/po/{year?}/{key?}/{SD?}/{flash?}/{bidder_id?}/{POno?}' , 'ExcelController@purhcaseOrder')->name('.pr');
        Route::get('/wfp/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@wfp')->name('.wfp');
        Route::get('/wfpConsolidated/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@wfpConsolidated')->name('.wfpConsolidated');
        Route::get('/ppmp/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@ppmp')->name('.ppmp');
        Route::get('/APPOffice/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@APPOffice')->name('.APPOffice');
        Route::get('/APPOfficeCategory/{year?}/{key?}/{SD?}/{flash?}' , 'ExcelController@APPOfficeCategory')->name('.APPOfficeCategory');
        Route::get('/DTRR/{year?}/{key?}/{flash?}' , 'ExcelController@DTRR')->name('.DTRR');
    });
});
