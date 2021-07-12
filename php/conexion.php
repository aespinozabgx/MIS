
<!DOCTYPE html>
<html lang="es" dir="ltr">
  <head>
    <meta charset="utf-8">
    <title></title>
    <style type="text/css">

      body
      {
        background: rgb(237, 237, 237, 1);
      }

      h2
      {

      }

      .contenedor
      {
        margin: 50px auto;
      }

    </style>
  </head>
  <body>
    <?php

    $serverName = "AXELEA7CFE\MSSQLSERVER02"; //serverName\instanceName

    // Since UID and PWD are not specified in the $connectionInfo array,
    // The connection will be attempted using Windows Authentication.
    $connectionInfo = array( "Database"=>"DESARROLLO", "CharacterSet" => "UTF-8");
    $conn = sqlsrv_connect( $serverName, $connectionInfo);

    if(!$conn)
    {
      echo "Connection could not be established.<br />";
      die( print_r( sqlsrv_errors(), true));
    }


      //$serverName = "tcp:192.168.200.144";
      // $connectionInfo = array(
      //     "Database" => "ADMINISTRACION_TI",
      //          "UID" => "sa",
      //          "PWD" => "C09%1_e4_",
      // "CharacterSet" => "UTF-8");
      //
      // $conn = sqlsrv_connect( $serverName, $connectionInfo );
      // if( $conn === false )
      // {
      //   echo "<h2>No se pudo conectar al servidor. Verifique su conexión</h2>";
      //   die();
      // }
      // else
      // {
      //   echo "Conectado a PRUEBAS";
      // }

    ?>
  </body>
</html>
