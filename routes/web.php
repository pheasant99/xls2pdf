<?php

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

/*
Route::get('/', function () {
    return view('welcome');
});
*/

Route::get('/', 'MainController@index');
Route::get('/2', 'MainController@index2');
Route::get('/3', 'MainController@index3');
Route::get('/DL', 'MainController@indexDL');
Route::get('/Test', 'MainController@indexTest');
Route::get('/DOBJ', 'MainController@indexDOBJ');
