<?php

namespace App\Http\Controllers;

use GuzzleHttp\Exception\ClientException;
use GuzzleHttp\Exception\GuzzleException;
use GuzzleHttp\Exception\RequestException;
use GuzzleHttp\Psr7\MultipartStream;
use Illuminate\Http\Request;
use Illuminate\Support\Carbon;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Redirect;
use Illuminate\Support\Facades\Response;
use Psr\Http\Message\ResponseInterface;
use function Illuminate\Support\Facades\Route;

class EfactController extends Controller
{
    /**
     * Display a listing of the resource.
     *
     * @return \Illuminate\Http\Response
     */
    public function Auth(Request $request)
    {

        sleep(5);

        $operacion = $request->input("ope");
        $tipdoc = $request->input("tipdoc");
        $client = new \GuzzleHttp\Client();

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        $urlprueba="https://ose-gw1.efact.pe:443/api-efact-ose/oauth/token";
        $url="https://ose.efact.pe/api-efact-ose/oauth/token";

        try {
            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", $url, [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {

                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                //' request.querystring (equivale)'
                $file = $request->input("file");
                //'redirect to (......)'
                return Redirect::route("sendfile", ["file" => $file, "token" => $token, "operacion" => $operacion, "tipdoc" => $tipdoc]);


            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            return response($e->getResponse()->getBody()->getContents(), 500, []);
            //return response($e->getMessage(), 500, []);
        }

    }

    public function sendFile(Request $request)
    {
        $client = new \GuzzleHttp\Client();
        $operacion = $request->input("operacion");
        $tipdoc = $request->input("tipdoc");

        $fullPath = "\\\\intranet2\\D\\EFACT\\daemon\\documents\\out\\$tipdoc\\" . $request->input("file");

        $ferr =  str_replace(".xml","","\\\\intranet2\\D\\EFACT\\daemon\\documents\\err\\$tipdoc\\" . $request->input("file")).".csv";

        while (!file_exists($fullPath)) {
            if(file_exists($ferr)){
                //unlink($ferr);
                //return "archivo generado csv no existe revise si hay error o intente nuevamente.";

            }
            sleep(1);
        }

        $bearer = "Bearer " . $request->input("token");

        try {
            //dd($bearer);

            $url      ="https://ose.efact.pe/api-efact-ose/v1/document";
            $urlprueba="https://ose-gw1.efact.pe:443/api-efact-ose/v1/document";

            $requestSendFile = $client->request('POST',$url ,
                [
                    'multipart' => [
                        [
                            'name' => 'file',
                            'contents' => fopen($fullPath, "r"),
                        ]
                    ],
                    'headers' => ["Authorization" => $bearer]
                ]
            );

            if ($requestSendFile->getStatusCode() >= 200 and $requestSendFile->getStatusCode() < 300) {
                $responseSendFileN = json_decode($requestSendFile->getBody(), true);
//                return $requestSendFile->getBody();
                $ticket = $responseSendFileN["description"];

                $url       = "https://ose.efact.pe/api-efact-ose/v1/cdr/" . $ticket . "";
                $urlprueba = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/cdr/" . $ticket . "";


                $requestCDR = $client->request("GET", $url, [
                    'headers' => ["Authorization" => $bearer]
                ]);


                $responseCDR = json_decode($requestCDR->getBody(), true);


                echo json_encode(["file" => $responseSendFileN, "cdr" => $responseCDR, "msg" => "espere unos segundos"]);
                DB::table("REIMPRIME")->insert(['OPERACION' => $operacion, 'TICKET' => $ticket]);
                sleep(15);
                return Redirect::route("show", ["tipo" => "pdf", "ticket" => "$ticket"]);
            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {
            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);

            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
        }
    }

    public function baja(Request $request)
    {
        $client = new \GuzzleHttp\Client();
        $operacion = $request->input("ope");

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        try {



            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", "https://ose.efact.pe/api-efact-ose/oauth/token", [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {
                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                $fullPath = "\\\\intranet2\\D\\EFACT\\daemon\\documents\\in\\" . $request->input("file");


                while (!file_exists($fullPath)) {
                    //echo \response()->json(["error" => "archivo generado csv no existe revise si hay error o intente nuevamente."]);
                    sleep(1);
                }

                $bearer = "Bearer " . $token;

                try {
                    //dd($bearer);
                    $url      ="https://ose.efact.pe/api-efact-ose/v1/document";
                    $urlprueba="https://ose-gw1.efact.pe:443/api-efact-ose/v1/document";
                    $requestSendFile = $client->request('POST', $url,
                        [
                            'multipart' => [
                                [
                                    'name' => 'file',
                                    'contents' => fopen($fullPath, "r"),
                                ]
                            ],
                            'headers' => ["Authorization" => $bearer]
                        ]
                    );

                    if ($requestSendFile->getStatusCode() >= 200 and $requestSendFile->getStatusCode() < 300) {
                        $responseSendFileN = json_decode($requestSendFile->getBody(), true);
//                return $requestSendFile->getBody();
                        $ticket = $responseSendFileN["description"];

                        DB::table("REIMPRIME")->insert(['OPERACION' => $operacion, 'TICKET' => $ticket]);
                        DB::table("efact_bajas")->insert(['operacion' => $operacion, 'TICKET' => $ticket,'fecbaja'=>Carbon::now()->format('d-m-Y H:i:s'),'estado'=>$requestSendFile->getStatusCode().""]);
                        sleep(10);
                        return $responseSendFileN;
                    }
                } catch (\GuzzleHttp\Exception\RequestException $e) {

                    //return response($e->getResponse()->getBody()->getContents(), 500, []);
                    $err  = $e->getResponse()->getBody()->getContents();
                    $rerr = json_decode($err, true);

                    return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
                }
            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            //return response($e->getResponse()->getBody()->getContents(), 500, []);
            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);

            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
            //return response($e->getMessage(), 500, []);
        }
    }

    public function download(Request $request)
    {

        //request = querystring
        $ticket = $request->input("ticket");
        $tipo = $request->input("tipo");
        $tipo2 = $tipo;
        if ($tipo != "pdf") {
            $tipo2 = "xml";
        }
//        $url = "";
//        switch ($tipo) {
//            case "PDF":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/pdf/$ticket";
//                break;
//            case "CDR":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/cdr/$ticket";
//                break;
//            case "XML":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/xml/$ticket";
//                break;
//        }
        //' solo estamos trabajando con el PDF, por eso se comenta el switch de arriba'
        // https://ose.efact.pe/api-efact-ose/v1/xml/{ticket}
        $url = "https://ose.efact.pe/api-efact-ose/v1/$tipo/$ticket";

        $client = new \GuzzleHttp\Client();

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        try {
            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", "https://ose.efact.pe/api-efact-ose/oauth/token", [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {
                $xxfile =  storage_path()."\\$ticket.$tipo2";
                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                $bearer = "Bearer " . $token;
                $requestDownload = $client->request("GET", "$url", [
                    'headers' => [
                        'Content-Type' => $contentType,
                        'Authorization' => $bearer
                    ],
                    'save_to' =>$xxfile,
//                    "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
                ]);
                //return  ($bearer);

                return Response::download($xxfile)->deleteFileAfterSend(true);


            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);

            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
        }
    }


    public function show(Request $request)
    {

        //request = querystring
        $ticket = $request->input("ticket");
        $download = $request->input("download");
        $tipo = $request->input("tipo");
        $tipo2 = $tipo;
        if ($tipo != "pdf") {
            $tipo2 = "xml";
        }
//        $url = "";
//        switch ($tipo) {
//            case "PDF":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/pdf/$ticket";
//                break;
//            case "CDR":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/cdr/$ticket";
//                break;
//            case "XML":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/xml/$ticket";
//                break;
//        }
        //' solo estamos trabajando con el PDF, por eso se comenta el switch de arriba'
        $url = "https://ose.efact.pe/api-efact-ose/v1/$tipo/$ticket";

        $client = new \GuzzleHttp\Client();

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        try {
            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", "https://ose.efact.pe/api-efact-ose/oauth/token", [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {

                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                $bearer = "Bearer " . $token;
                $xxfile =  storage_path()."\\$ticket.$tipo2";
                $requestDownload = $client->request("GET", "$url", [
                    'headers' => [
                        'Content-Type' => $contentType,
                        'Authorization' => $bearer
                    ],
                    'save_to' =>$xxfile,
//                    "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
                ]);
                //return  ($bearer);
                if(!empty($download)){
                    return Response::download($xxfile);
                }
                return Response::file($xxfile)->deleteFileAfterSend(true);


            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            //return response($e->getResponse()->getBody()->getContents(), 500, []);
            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);
            if(empty($rerr["code"])){
                $rercode='500';
            }else{
                $rercode = $rerr["code"];
            }
            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rercode."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
        }
    }


    public function consultaTicket(Request $request)
    {

        //request = querystring
        $ruc = "20600689101";
        $documento = $request->input("documento");
        $tipdoc = $request->input("tipdoc");

//        $url = "";
//        switch ($tipo) {
//            case "PDF":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/pdf/$ticket";
//                break;
//            case "CDR":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/cdr/$ticket";
//                break;
//            case "XML":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/xml/$ticket";
//                break;
//        }
        //' solo estamos trabajando con el PDF, por eso se comenta el switch de arriba'
        $url = "https://ose.efact.pe/api-efact-ose/v1/ticket/$ruc-$tipdoc-$documento";
//        return $url;
        $client = new \GuzzleHttp\Client();

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        try {
            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", "https://ose.efact.pe/api-efact-ose/oauth/token", [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {

                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                $bearer = "Bearer " . $token;

                $requestDownload = $client->request("GET", "$url", [
                    'headers' => [
                        'Content-Type' => $contentType,
                        'Authorization' => $bearer
                    ],
//                    'save_to' => "\\\\intranet2\\D\\ventas_new\\efact_2\\$ticket.$tipo2",
//                    "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
                ]);
                //return  ($bearer);
                if ($requestDownload->getStatusCode() >= 200 && $requestDownload->getStatusCode() <= 299) {
                    return $requestDownload->getBody()->getContents();
                }
                throw new RequestException($requestDownload);
            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            //return response($e->getResponse()->getBody()->getContents(), 500, []);
            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);

            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
        }
    }

    public function consultaTicketYGraba(Request $request)
    {

        //request = querystring
        $ruc = "20600689101";
        $documento = $request->input("documento");
        $operacion = $request->input("operacion");
        $tipdoc = $request->input("tipdoc");

//        $url = "";
//        switch ($tipo) {
//            case "PDF":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/pdf/$ticket";
//                break;
//            case "CDR":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/cdr/$ticket";
//                break;
//            case "XML":
//                $url = "https://ose-gw1.efact.pe:443/api-efact-ose/v1/xml/$ticket";
//                break;
//        }
        //' solo estamos trabajando con el PDF, por eso se comenta el switch de arriba'
        $url = "https://ose.efact.pe/api-efact-ose/v1/ticket/$ruc-$tipdoc-$documento";
//        return $url;
        $client = new \GuzzleHttp\Client();

        $user = "20600689101";
        $password = "f81a138412a7df5a950ed12085ae8eb012676be6c22b4a5e4172981c43fff2d8";
        $credential64 = "Basic " . base64_encode(utf8_encode("client:secret"));
        $contentType = ['content-type' => 'application/x-www-form-urlencoded'];

        try {
            //envio la peticion http con headers,auth, parametros del formulario
            $requestAuth = $client->request("POST", "https://ose.efact.pe/api-efact-ose/oauth/token", [
                "auth" => ["client", "secret"],
                'headers' => [
                    'Content-Type' => $contentType,
                    'Authorization' => $credential64
                ],
                "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
            ]);

            //valido que el estado de la pagina sea de 200 hasta 299 para poder capturar el token
            if ($requestAuth->getStatusCode() >= 200 && $requestAuth->getStatusCode() < 300) {

                $response = json_decode($requestAuth->getBody(), true);
                $token = $response["access_token"];
                $bearer = "Bearer " . $token;

                $requestDownload = $client->request("GET", "$url", [
                    'headers' => [
                        'Content-Type' => $contentType,
                        'Authorization' => $bearer
                    ],
//                    'save_to' => "\\\\intranet2\\D\\ventas_new\\efact_2\\$ticket.$tipo2",
//                    "form_params" => ["username" => $user, "password" => $password, "grant_type" => "password"]
                ]);
                //return  ($bearer);
                if ($requestDownload->getStatusCode() >= 200 && $requestDownload->getStatusCode() <= 299) {
                    $result = json_decode($requestDownload->getBody()->getContents(), true);
                    $tt = $result["tickets"][0];
                    DB::table("REIMPRIME")->where('OPERACION', $operacion)->delete();
                    DB::table("REIMPRIME")->insert(['OPERACION' => $operacion, 'TICKET' => $tt]);
                    return "<script>alert('se actualizo correctamente');</script><script>window.close()</script>";
                }
                throw new RequestException($requestDownload);
            }
        } catch (\GuzzleHttp\Exception\RequestException $e) {

            //return response($e->getResponse()->getBody()->getContents(), 500, []);
            $err  = $e->getResponse()->getBody()->getContents();
            $rerr = json_decode($err, true);

            return response("<table style='width:100%'><tr><th>Codigo</th><th>Mensaje</th></tr><tr><td>".$rerr["code"]."</td><td>".$rerr["description"]."</td></tr></table><script>alert('".$rerr["description"]."')</script>", 202, []);
        }
    }
}
