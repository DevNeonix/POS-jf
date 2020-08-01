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

use App\Movimcab;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return view('welcome');
});

Route::get('operaciones', function (Builder $builder) {
    if (request()->ajax()) {
        return DataTables::of(Movimcab::query())->toJson();
    }

    $html = $builder->columns([
        ['data' => 'OPERACION'],
        ['data' => 'TIENDA'],
        ['data' => 'CODDOC'],
        ['data' => 'SERIE'],
        ['data' => 'NUMDOC']]);

    return view('movimientos', compact('html'));
});

//EFACT
Route::get('auth', 'EfactController@auth');

Route::get('authweb', 'EfactController@authweb');

Route::get('baja', 'EfactController@baja');
Route::get('sendfile', 'EfactController@sendFile')->name("sendfile");
Route::get('download', 'EfactController@download')->name("download");
Route::get('show', 'EfactController@show')->name("show");
Route::get('consulta', 'EfactController@consultaTicket')->name("consulta");
Route::get('grabaticket', 'EfactController@consultaTicketYGraba')->name("consultagraba");
