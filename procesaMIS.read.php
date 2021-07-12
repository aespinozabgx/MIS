<?php

require 'vendor/autoload.php';
require 'php/conexion.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


  if (isset($_POST['cargaDatos']))
  {

    echo "<br><h3>Month to update: " .  $_POST['daterange'] . "</h3><br>";

    // STAR OF REPORT MATRIZ

      $T_repIntProv    = Array();  // Matriz de reporte de intereses y provisiones
      $T_repMorosidad  = Array();  // Matriz de reporte de morosidad
      $T_baseProyectos = Array();  // Matriz de proyectos base (DB)

    // END OF REPORT MATRIX


    // READ REPORTE MOROSIDAD
    if($_FILES["reporteMorosidad"]["name"] != '')
    {

        $allowed_extension = array('xls', 'csv', 'xlsx');
        $file_array = explode(".", $_FILES["reporteMorosidad"]["name"]);
        $file_extension = end($file_array);

        if(in_array($file_extension, $allowed_extension))
        {

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

                $T_repMorosidad[$indexVector][0] = ($data[$x][6]);
                $T_repMorosidad[$indexVector][1] = ($data[$x][14]);
                $indexVector++;
              }


            $message = '<div class="alert alert-success">Data Imported Successfully</div>';

            if (isset($output))
            {
              echo $output;
            }

          } // END RECORRE HOJAS

          $message = '<div class="alert alert-success">Data Imported Successfully</div>';

        }
        else
        {
          $message = '<div class="alert alert-danger">Only .xls .csv or .xlsx file allowed</div>';
        }
    }
    else
    {
      $message = '<div class="alert alert-danger">Please Select File</div>';
    }
    echo $message;
    // END OF READ REPORTE MOROSIDAD

    echo "<br><hr><br>";

    // READ REPORTE PROVISIONES
    if($_FILES["repIntProv"]["name"] != '')
    {
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

    		$message = '<div class="alert alert-success">Data Imported Successfully</div>';

    	}
    	else
    	{
    		$message = '<div class="alert alert-danger">Only .xls .csv or .xlsx file allowed</div>';
    	}
    }
    else
    {
    	$message = '<div class="alert alert-danger">Please Select File</div>';
    }

    echo $message;

    // echo "<pre>";
    // var_dump($matrizRepIntProv[0]);
    // echo "</pre>";

    // echo "<pre>";
    // print_r($matrizRepIntProv[0]);
    // echo "</pre>";

    //echo print_r($matrizRepIntProv[0]['6']["C"], true);



  	$p = 0; /* Sólo para el contador */

    // Tomo los datos de el array de arrays y valido la fecha del ultimo pago
    // con la fecha que se pretende actualizar.


    /*
      Validamos si existe pago de cada proyecto realizados en el periodo que se pretende actualizar
    */
    $remover = ["$", ",", " "];
  	foreach ($matrizRepIntProv as $vals)
  	{
        //echo "Eval: " . $vals['9']["B"] . "<br>";
  			if (isset($vals['9']["B"]) && date('m-Y', strtotime($vals['9']["B"])) == $_POST['daterange'])
  			{

          //echo "SI: " . date('m-Y', strtotime($vals['9']["B"])) . " - " . $_POST['daterange'] . "<br>";
  				$T_repIntProv[$p][0] = $vals['6']["C"];
  				$T_repIntProv[$p][1] = $vals['9']["B"];
  				$T_repIntProv[$p][2] = (str_replace($remover, "", $vals['9']["C"])) + 0.0;
  				$p++;
  			}
        else
        {
          $T_repIntProv[$p][0] = $vals['6']["C"];
          $T_repIntProv[$p][1] = "";
          $T_repIntProv[$p][2] = 0;
          $p++;
        }
  	}

    // Ordenar ASC
    asort($T_repIntProv);


    // Show REP INT PROV
    foreach ($T_repIntProv as $key => $value)
    {
    	//echo $key . " - " . $value[0] . " - " . (!empty($value[1]) ? date("d-m-Y", strtotime($value[1])) : "" ) . " - ". $value[2] . "<br><br>";
    }



    // END OF READ REPORTE PROVISIONES


    // START OF MERGE DATA

    $T_repIntProv   = array_values($T_repIntProv);
    $T_repMorosidad = array_values($T_repMorosidad);


    // Relaciono REP DE MOROSIDAD CON REPORTE DE INT Y PROVISIONES


    /* Set up and execute the query. */
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
    sqlsrv_free_stmt( $stmt);
    sqlsrv_close( $conn);

    for ($i=0; $i < sizeof($T_baseProyectos); $i++)
    {

        echo $T_baseProyectos[$i][0] . "<br>";

    }

    $contador = 1;
    for ($col=0; $col < sizeof($T_repMorosidad); $col++)
    {
      for ($row=0; $row < sizeof($T_repIntProv); $row++)
      {
        if ($T_repIntProv[$row][0] == $T_repMorosidad[$col][0])
        {
            echo $contador . ".- " . $T_repIntProv[$row][0] . "<br>";
            echo $T_repIntProv[$row][2] . " - " . $T_repMorosidad[$col][1] . "<br><br>";
            $contador++;
        } else {
          //echo $T_repIntProv[$row][0] . " - " . $T_repMorosidad[$col][0];
        }
      }
    }
    // Relaciono REP DE MOROSIDAD CON REPORTE DE INT Y PROVISIONES

    //

    //

    // END OF MERGE DATA

} // END OF CARGA DE DATOS






?>
