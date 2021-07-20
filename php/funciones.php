<?php

function validaCargaColateral($conn, $mesActualziar)
{
    // Esta madre revisa si existe información previa cargada del mes que se intenta actualizar

    /* Obtener mes */
    $fechaPorciones = explode(' ',$mesActualziar);
    $onlyDate = $fechaPorciones[0];
    $onlyDate = explode('-',$mesActualziar);
    $mes = $onlyDate[1];
    /* Fin obtener mes */


    $tsql = "SELECT * FROM MIS_colaterales WHERE datepart(month, FECH_COLATERAL)= ?";
    $params = array($mes);
    $stmt = sqlsrv_query($conn, $tsql, $params, array("Scrollable" => SQLSRV_CURSOR_KEYSET));

    $row_count = sqlsrv_num_rows( $stmt );

    if ($row_count === false)
       return false;
    else if ($row_count >0)
       return true;

}

function fechaCastellano($fecha)
{
  $fecha = substr($fecha, 0, 10);
  $numeroDia = date('d', strtotime($fecha));
  $dia = date('l', strtotime($fecha));
  $mes = date('F', strtotime($fecha));
  $anio = date('Y', strtotime($fecha));
  $dias_ES = array("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo");
  $dias_EN = array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday");
  $nombredia = str_replace($dias_EN, $dias_ES, $dia);
  $meses_ES = array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre");
  $meses_EN = array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
  $nombreMes = str_replace($meses_EN, $meses_ES, $mes);
  return $nombreMes."-".$anio;
}


function fromExcelToLinux($excel_time)
{
    return ($excel_time-25569)*86400;
}

function in_array_r($item , $array)
{
  /*
    Checks if exists in multidimensional arrays
  */
    return preg_match('/"'.preg_quote($item, '/').'"/i' , json_encode($array));
}

// START  READ REPORTE SHF
function readReporteSHF($conn)
{

    if($_FILES["reporteSHF"]["name"] != '')
    {

        $allowed_extension = array('xls', 'csv', 'xlsx');
        $file_array = explode(".", $_FILES["reporteSHF"]["name"]);
        $file_extension = end($file_array);

        if(in_array($file_extension, $allowed_extension))
        {
            $T_aux = Array();
            $file_name = time() . '.' . $file_extension;
            move_uploaded_file($_FILES['reporteSHF']['tmp_name'], $file_name);
            $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_name);
            $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);
            $reader->setReadDataOnly(true);
            $spreadsheet = $reader->load($file_name);

            unlink($file_name);

            $data = $spreadsheet->getActiveSheet()->toArray();

            $sheetCount = 1;
            //$sheetCount = $spreadsheet->getSheetCount();



            // START RECORRE HOJAS
            for ($z = 0; $z < $sheetCount; $z++)
            {

                $sheet = $spreadsheet->getSheet($z);
                $matriz = $sheet->toArray(null, true, true, true);

                //echo "<br>Cols 	X: " . max(array_map('count', $sheetData)) . ",
                //				Rows  Y: " . sizeof($sheetData) . "<br><hr><br>";


            } // END RECORRE HOJAS

            $titulos = Array();

            foreach ($matriz[1] as $key => $value) { $titulos[] = $value; }
            unset($matriz[1]);

            $matriz = array_values($matriz);


            $message = '<div class="alert alert-success">Información importada correctamente (SHF)</div>';

        }
        else
        {
          $message = '<div class="alert alert-danger">Sólo se permiten .xls .csv o .xlsx</div>';
        }

    }
    else
    {
      $message = '<div class="alert alert-danger">Seleccione un documento.</div>';
    }

    echo $message;



    foreach($matriz as $row)
    {

        $linux_time = fromExcelToLinux($row['B']);
        $tsql = "INSERT INTO MIS_temp_shf (NOM_CONJUNTO, FECH_FIN_CONTRATO, AO_VIV_ACTIVAS, VIV_LIB_PERIODO, MONTO_MIN_EN_EL_PERIODO, MONTO_AMORT_EN_EL_PERIODO) VALUES (?, ?, ?, ?, ?, ?)";
        $params = Array($row['A'], date("Y-m-d", $linux_time), $row['C'], $row['D'], $row['E'], $row['F']);
        $stmt = sqlsrv_query($conn, $tsql, $params);

        if( !$stmt)
        {
            if( ($errors = sqlsrv_errors() ) != null)
            {
                foreach( $errors as $error )
                {
                    echo "SQLSTATE: ".$error[ 'SQLSTATE']."<br />";
                    echo "code: ".$error[ 'code']."<br />";
                    echo "message: ".$error[ 'message']."<br />";
                }
            }
        }


    }


}
// ENF OF READ REPORTE SHF

// START  READ TABLA INTERMEDIA
function readTablaIntermedia($conn)
{

    $tsql = "SELECT * FROM MIS_TABLA_INTERMEDIA";
    $stmt = sqlsrv_query( $conn, $tsql);
    if( $stmt === false)
    {
         echo "Error in query preparation/execution.\n";
         die( print_r( sqlsrv_errors(), true));
    }

    $i = 0; /* Para contador only*/
    while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC))
    {
        $matriz[$i][0] = $row['UNO'];
        $matriz[$i][1] = $row['DOS'];
        $i++;
    }

    /* Free statement and connection resources. */
    // sqlsrv_free_stmt( $stmt);
    // sqlsrv_close( $conn);

    return $matriz;

}
// ENF OF READ TABLA INTERMEDIA

// READ PROJECTS FROM DB
function readProyectosDB($conn)
{

    $tsql = "SELECT NOM_CONJUNTO, NOM_PROMOTOR FROM MIS_proyectos";
    $stmt = sqlsrv_query( $conn, $tsql);
    if( $stmt === false)
    {
         echo "Error in query preparation/execution.\n";
         die( print_r( sqlsrv_errors(), true));
    }

    $r = 0; /* Para contador only*/
    while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC))
    {
        $T_baseProyectos[$r][0] = $row['NOM_CONJUNTO'];
        $T_baseProyectos[$r][1] = $row['NOM_PROMOTOR'];
        $r++;
    }

    /* Free statement and connection resources. */
    // sqlsrv_free_stmt( $stmt);
    // sqlsrv_close( $conn);

    if ($r > 0)
    {
      return $T_baseProyectos;
    }
    else
    {
      return false;
    }

}
// END OF READ PROJECTS DB


// READ REPORTE MOROSIDAD
function readReporteMorosidad($conn)
{

  if($_FILES["reporteMorosidad"]["name"] != '')
  {

      $allowed_extension = array('xls', 'csv', 'xlsx');
      $file_array = explode(".", $_FILES["reporteMorosidad"]["name"]);
      $file_extension = end($file_array);

      if(in_array($file_extension, $allowed_extension))
      {
          $T_aux = Array();
          $file_name = time() . '.' . $file_extension;
          move_uploaded_file($_FILES['reporteMorosidad']['tmp_name'], $file_name);
          $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_name);
          $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);
          $reader->setReadDataOnly(true);
          $spreadsheet = $reader->load($file_name);

          unlink($file_name);

          $data = $spreadsheet->getActiveSheet()->toArray();

          $sheetCount = 1;
          //$sheetCount = $spreadsheet->getSheetCount();


          // echo "<pre>";
          // var_dump($data);
          // echo "</pre>";

          // START RECORRE HOJAS
          for ($z = 0; $z < $sheetCount; $z++)
          {

              $sheet = $spreadsheet->getSheet($z);
              $sheetData = $sheet->toArray(null, true, true, true);

              //echo "<br>Cols 	X: " . max(array_map('count', $sheetData)) . ",
              //				Rows  Y: " . sizeof($sheetData) . "<br><hr><br>";



              for ($x=2, $cont = 1, $indexVector=0; $x < sizeof($sheetData); $x++, $cont++)
              {
                // Tomo los valores fijos con las columnas del archivo

                //$T_aux[$indexVector]['NOM_CONJUNTO'] = ($data[$x][6]);
                //$T_aux[$indexVector]['INTERESES_RM'] = ($data[$x][15]);
                $T_aux[$indexVector][0] = ($data[$x][6]);
                $T_aux[$indexVector][1] = ($data[$x][15]);
                $indexVector++;
              }


          } // END RECORRE HOJAS

          $message = '<div class="alert alert-success">Información importada correctamente (Morosidad)</div>';

      }
      else
      {
        $message = '<div class="alert alert-danger">Sólo se permiten .xls .csv o .xlsx</div>';
      }

  }
  else
  {
    $message = '<div class="alert alert-danger">Seleccione un documento.</div>';
  }

  echo $message;

  $T_aux = array_values($T_aux);

  //return $T_aux;

  foreach($T_aux as $row)
  {

      $tsql = "INSERT INTO MIS_temp_morosidad (PROYECTO, INT_DEVENGADO) VALUES (?,?)";
      $params = Array($row[0], $row[1]);
      $stmt = sqlsrv_query($conn, $tsql, $params);

  }

}
// END OF READ REPORTE MOROSIDAD


// READ REPORTE INTERESES Y PROVISIONES
function readReporteIntProv($conn)
{

    if($_FILES["repIntProv"]["name"] != '')
    {
      $T_aux = Array();
      $allowed_extension = array('xls', 'csv', 'xlsx');
      $file_array = explode(".", $_FILES["repIntProv"]["name"]);
      $file_extension = end($file_array);

      if(in_array($file_extension, $allowed_extension))
      {
        $file_name = time() . '.' . $file_extension;
        move_uploaded_file($_FILES['repIntProv']['tmp_name'], $file_name);
        $file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_name);
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);

        $spreadsheet = $reader->load($file_name);



        unlink($file_name);

        $data = $spreadsheet->getActiveSheet()->toArray();
        $sheetCount = $spreadsheet->getSheetCount();

        $matrizRepIntProv = Array();

        for ($z = 0; $z < $sheetCount; $z++)
        {

            $sheet = $spreadsheet->getSheet($z);
            $matrizRepIntProv[$z] = $sheet->toArray(null, true, true, true);

        }

        $message = '<div class="alert alert-success">Información importada correctamente (Int y prov)</div>';

      }
      else
      {

        $message = '<div class="alert alert-danger">Sólo se permiten .xls .csv o .xlsx</div>';

      }

    }
    else
    {

      $message = '<div class="alert alert-danger">Seleccione un documento.</div>';

    }

    // echo "<pre>";
    // var_dump($matrizRepIntProv[0]);
    // echo "</pre>";

    // echo "Matriz_RepIntProv: <br><br>";
    // echo "<pre>";
    // print_r($matrizRepIntProv[0]);
    // echo "</pre>";

    //echo print_r($matrizRepIntProv[0]['6']["C"], true);

    $p = 0; /* Sólo para el contador */

    /*
      Validamos si existe pago de cada proyecto realizados en el periodo que se pretende actualizar
    */
    $remover = ["$", ",", " "];
    foreach ($matrizRepIntProv as $vals)
    {
        //echo "Eval: " . $vals['9']["B"] . "<br>";
        if (isset($vals['9']["B"]) && date('m-Y', strtotime($vals['9']["B"])) == $_POST['toUpdate'])
        {

          //echo "SI: " . date('m-Y', strtotime($vals['9']["B"])) . " - " . $_POST['daterange'] . "<br>";
          $T_aux[$p][0] = $vals['6']["C"];
          $T_aux[$p][1] = $vals['9']["B"];
          $T_aux[$p][2] = (str_replace($remover, "", $vals['9']["C"])) + 0.0;
          $p++;
        }
        else
        {
          $T_aux[$p][0] = $vals['6']["C"];
          $T_aux[$p][1] = "";
          $T_aux[$p][2] = 0;
          $p++;
        }

    }

    // Ordenar ASC
    asort($T_aux);

    echo $message;

    // return matrix
    $T_aux = array_values($T_aux);

    foreach($T_aux as $row)
    {

        $tsql = "INSERT INTO MIS_temp_intprov (PROYECTO, INTERESES) VALUES (?,?)";
        $params = Array($row[0], $row[2]);
        $stmt = sqlsrv_query($conn, $tsql, $params);

    }

}
// END OF READ REPORTE INTERESES PROVISIONES

// Formatear moneda
function formatoMoneda($val,$symbol='$',$r=2)
{

  $n = $val;
      $c = is_float($n) ? 1 : number_format($n,$r);
      $d = '.';
      $t = ',';
      $sign = ($n < 0) ? '-' : '';
      $i = $n=number_format(abs($n),$r);
      $j = (($j = strlen($i)) > 3) ? $j % 3 : 0;

     return  $symbol.$sign .($j ? substr($i,0, $j) + $t : '').preg_replace('/(\d{3})(?=\d)/',"$1" + $t,substr($i,$j)) ;

}
// Fin Formatear moneda


function limpia ($dato)
{
  $dato = trim($dato);
  $dato = stripslashes($dato);
  $dato = strip_tags($dato);
  //$dato = htmlspecialchars($dato);
  //$dato = htmlentities($dato);

  return $dato;
}


?>
