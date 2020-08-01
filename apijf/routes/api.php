<?php

use App\Movimcab;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
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

Route::post('auth', function (Request $request) {


    $u = collect(DB::select("select rtrim(usuario) usuario,rtrim(nombres) nombres from usuarios where usuario = ? and clave = ?",
        [$request->input("usuario"), $request->input("password")]
    ))->first();

    return response()->json($u);
});
