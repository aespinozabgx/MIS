<?php

require 'vendor/autoload.php';
require 'php/funciones.php';

require 'php/conexion.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


    if (isset($_POST['btnFiltra']))
    {
      header('Location: index.php?date=' . $_POST['fechaReporte']);
    }

    if (isset($_POST['cargaDatos']))
    {
        $mesActualziar = $_POST['daterange'];
        $mesActualizarDate = "01-" . $_POST['daterange'];
        $mesActualziar=explode('-',$mesActualziar);
        $mes = (int)$mesActualziar[0] + 0;
        $mes -= 1;
        $mesPrevio = $mes;
        $mesPrevio = "01-" . sprintf("%02d", $mesPrevio) . "-" . $mesActualziar[1];
        $mesPrevio = date("d-m-Y",strtotime($mesPrevio));

        echo "<br>Month to update: " .  $mesActualizarDate . "<br>";
        echo "<br>Previous Month: " .  $mesPrevio . "<br>";




        // STAR OF REPORT MATRIZ

        $T_proyectos     = Array();  // Matriz de proyectos base (DB)
        $T_repIntProv    = Array();  // Matriz de reporte de intereses y provisiones
        $T_repMorosidad  = Array();  // Matriz de reporte de morosidad
        $T_SHF           = Array();  // Matriz de datos SHF
        $T_intermedia    = Array();  // Matriz de tabla intermedia

        // END OF REPORT MATRIX


        readReporteSHF($conn);
        readReporteMorosidad($conn);
        readReporteIntProv($conn);


        //$T_proyectos     = readProyectosDB($conn);
        //$T_intermedia    = readTablaIntermedia($conn);

        echo "<br><br>";

        /* PASO 5 */
        // $SQL = "SELECT

          // MIS_CAT_proyectos.COLATERAL,
          // MIS_colaterales.NOM_PROYECTO,
          // MIS_colaterales.FECH_COLATERAL,
          // MIS_CAT_proyectos.CVE_CRE_IF,
          // MIS_CAT_proyectos.CVE_CRE_ID_OFERTA,
          // MIS_CAT_proyectos.NOM_PROYECTO,
          // MIS_CAT_proyectos.NUM_REF_SHF,
          // MIS_CAT_proyectos.NOM_PROMOTOR,
          // MIS_CAT_proyectos.TIPO_CREDITO,
          // MIS_CAT_proyectos.UBICACIÓN_EDO,
          // MIS_CAT_proyectos.UBICACIÓN_MUN,
          // MIS_CAT_proyectos.FECH_INI_CONTRATO,
          // MIS_CAT_proyectos.LINEA_DE_CRE_POR_PROYECTO,
          // MIS_CAT_proyectos.VALOR_PROYECTO,
          // MIS_CAT_proyectos.TASA_INTERES,
          // MIS_CAT_proyectos.VIV_TOTALES_PROYECTO,
          //
          // MIS_colaterales.FECH_FIN_CONTRATO,
          // MIS_colaterales.AO_VIV_ACTIVAS,
          // MIS_colaterales.VIV_LIB_PERIODO,
          // MIS_colaterales.MONTO_MIN_PERIODO,
          // MIS_colaterales.MONTO_AMORTIZADO_PERIODO,
          //
          // MIS_colaterales.MONTO_AMORT_ACUM_FIN_P,
          // MIS_colaterales.ACUM_VIV_LIB_FIN_P,
          // MIS_colaterales.MONTO_MIN_ACUM_FIN_P,
          // MIS_colaterales.MONTO_POR_DISPONER,
          // MIS_colaterales.SALDO_INS_CARTERA_FIN_P,
          // MIS_colaterales.VIV_LIB_CORTE_ANTERIOR,
          // MIS_colaterales.MONTO_AMORT_ACUM_P_ANTERIOR,
          // MIS_colaterales.SALDO_INS_P_ANTERIOR,
          // MIS_colaterales.MONTO_MIN_ACUM_P_ANTERIOR,
          // MIS_colaterales.INT_COBRADOS_PERIODO,
          // MIS_colaterales.NUM_MESES_MOROSOS,
          // MIS_colaterales.MONTO_INT_DEV_NO_CUBIERTOS
          // FROM MIS_CAT_proyectos INNER JOIN MIS_colaterales ON MIS_CAT_proyectos.NOM_PROYECTO = MIS_colaterales.NOM_PROYECTO";

          /* ===================================== */

        $tsql = "SELECT
        MIS_CAT_proyectos.COLATERAL,
        MIS_CAT_proyectos.CVE_CRE_IF,
        MIS_CAT_proyectos.CVE_CRE_ID_OFERTA,
        MIS_CAT_proyectos.NUM_REF_SHF,
        MIS_CAT_proyectos.NOM_PROYECTO,
        MIS_CAT_proyectos.NOM_PROMOTOR,
        MIS_CAT_proyectos.TIPO_CREDITO,
        MIS_CAT_proyectos.UBICACIÓN_EDO,
        MIS_CAT_proyectos.UBICACIÓN_MUN,
        MIS_CAT_proyectos.FECH_INI_CONTRATO,
        MIS_temp_shf.FECH_FIN_CONTRATO,
        MIS_CAT_proyectos.LINEA_DE_CRE_POR_PROYECTO,
        MIS_CAT_proyectos.VALOR_PROYECTO,
        MIS_temp_shf.AO_VIV_ACTIVAS,
        MIS_CAT_proyectos.TASA_INTERES,
        MIS_CAT_proyectos.VIV_TOTALES_PROYECTO,
        /* Viviendas Liberadas al Corte Anterior  */
        MIS_temp_shf.VIV_LIB_PERIODO,
        /* acum viv lib a fin periodo */
        /* monto min acum periodo ant */
        MIS_temp_shf.MONTO_MIN_EN_EL_PERIODO,
        /* monto min acum fin periodo */
        /* monto por disponer */
        /* monto amort acum periodo ant */
        MIS_temp_shf.MONTO_AMORT_EN_EL_PERIODO,
        /* monto amort acum fin periodo */
        /* saldo ins periodo ant */
        /* saldo ins cartera fin periodo */
        MIS_temp_intprov.INTERESES AS REP_INTPROV_INTERESES_DEV_NO_CUBIERTOS,
        /* comisiones cobradas */
        /* meses morosos */
        MIS_temp_morosidad.INT_DEVENGADO AS REP_MOR_INTERESES
        FROM MIS_temp_morosidad
        INNER JOIN MIS_TABLA_INTERMEDIA ON MIS_TABLA_INTERMEDIA.DOS = MIS_temp_morosidad.PROYECTO
        INNER JOIN MIS_temp_intprov ON MIS_temp_intprov.PROYECTO = MIS_TABLA_INTERMEDIA.DOS
        INNER JOIN MIS_CAT_proyectos ON MIS_CAT_proyectos.NOM_PROYECTO = MIS_TABLA_INTERMEDIA.UNO
        INNER JOIN MIS_temp_shf ON MIS_temp_shf.NOM_CONJUNTO = MIS_TABLA_INTERMEDIA.UNO";

        $stmt = sqlsrv_query( $conn, $tsql);
        if( $stmt === false)
        {
             echo "Error in query preparation/execution.\n";
             die( print_r( sqlsrv_errors(), true));
        }

        /* Retrieve each row as an associative array and display the results.*/

        //echo "
        // <table class=\"table table-bordered text-dark\" id=\"dataTable\" width=\"100%\" cellspacing=\"0\"  data-order='[[ 2, \"desc\" ]]' data-page-length='25'>
        // <thead>
        //   <tr>
        //     <th>#</th>
        //     <th>CVE_CRED_IF</th>
        //     <th>CVE_CRED_ID_OFERTA</th>
        //     <th>NUM_REF_SHF</th>
        //     <th>NOM_CONJUNTO</th>
        //     <th>NOM_PROMOTOR</th>
        //     <th>TIPO_CREDITO</th>
        //     <th>UBICACION_ESTADO</th>
        //     <th>UBICACION_MUNICIPIO</th>
        //     <th>FECH_INI_CONTRATO</th>
        //     <th>LINEA_DE_CREDITO_POR_PROYECTO</th>
        //     <th>VALOR_PROYECTO</th>
        //     <th>TASA_INTERES</th>
        //     <th>VIVIENDAS_TOTALES_DEL_PROYECTO</th>
        //     <th>PROYECTO AS PROYECTO_MAY</th>
        //     <th>UNO AS PROYECTO_MIN</th>
        //     <th>INT_DEVENGADO AS REP_MOR_INTERESES</th>
        //     <th>INTERESES AS REP_INTPROV_INTERESES</th>
        //   </tr>
        // </thead>
        // <tfoot>
        //   <tr>
        //     <th>#</th>
        //     <th>CVE_CRED_IF</th>
        //     <th>CVE_CRED_ID_OFERTA</th>
        //     <th>NUM_REF_SHF</th>
        //     <th>NOM_CONJUNTO</th>
        //     <th>NOM_PROMOTOR</th>
        //     <th>TIPO_CREDITO</th>
        //     <th>UBICACION_ESTADO</th>
        //     <th>UBICACION_MUNICIPIO</th>
        //     <th>FECH_INI_CONTRATO</th>
        //     <th>LINEA_DE_CREDITO_POR_PROYECTO</th>
        //     <th>VALOR_PROYECTO</th>
        //     <th>TASA_INTERES</th>
        //     <th>VIVIENDAS_TOTALES_DEL_PROYECTO</th>
        //     <th>PROYECTO AS PROYECTO_MAY</th>
        //     <th>UNO AS PROYECTO_MIN</th>
        //     <th>INT_DEVENGADO AS REP_MOR_INTERESES</th>
        //     <th>INTERESES AS REP_INTPROV_INTERESES</th>
        //   </tr>
        // </tfoot>
        // <tbody>";

        $contador = 0;
        while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_NUMERIC))
        {
              // echo "<tr>";
              // echo "<td>" . $contador . "</td>";
              // echo "<td>" . $row['0'] . "</td>";
              // echo "<td>" . $row['1'] . "</td>";
              // echo "<td>" . $row['2'] . "</td>";
              // echo "<td>" . $row['3'] . "</td>";
              // echo "<td>" . $row['4'] . "</td>";
              // echo "<td>" . $row['5'] . "</td>";
              // echo "<td>" . $row['6'] . "</td>";
              // echo "<td>" . $row['7'] . "</td>";
              // echo "<td>" . date_format($row['8'], "d-m-Y") . "</td>";
              // echo "<td>" . $row['9'] . "</td>";
              // echo "<td>" . $row['10'] . "</td>";
              // echo "<td>" . $row['11'] . "</td>";
              // echo "<td>" . $row['12'] . "</td>";
              // echo "<td>" . $row['13'] . "</td>";
              // echo "<td>" . $row['14'] . "</td>";
              // echo "<td>" . $row['15'] . "</td>";
              // echo "<td>" . $row['16'] . "</td>";
              // echo "</tr>";
              $matriz[$contador][0]  = $row['0'];
              $matriz[$contador][1]  = $row['1'];
              $matriz[$contador][2]  = $row['2'];
              $matriz[$contador][3]  = $row['3'];
              $matriz[$contador][4]  = $row['4'];
              $matriz[$contador][5]  = $row['5'];
              $matriz[$contador][6]  = $row['6'];
              $matriz[$contador][7]  = $row['7'];
              $matriz[$contador][8]  = $row['8'];
              $matriz[$contador][9]  = $row['9'];
              $matriz[$contador][10] = $row['10'];
              $matriz[$contador][11] = $row['11'];
              $matriz[$contador][12] = $row['12'];
              $matriz[$contador][13] = $row['13'];
              $matriz[$contador][14] = $row['14'];
              $matriz[$contador][15] = $row['15'];
              $matriz[$contador][16] = $row['16'];

              /* shf */
              $matriz[$contador][17] = $row['17'];
              $matriz[$contador][18] = $row['18'];
              $matriz[$contador][19] = $row['19'];
              $matriz[$contador][20] = $row['20'];

              $contador++;
        }


        //echo '</tbody></table>';

        /* Elimino tablas temporales */
        $tsql = "DELETE FROM MIS_temp_intprov; DELETE FROM MIS_temp_morosidad; DELETE FROM MIS_temp_shf";

        /* Set parameter values. */
        $params = array();

        /* Prepare and execute the query. */
        $stmt = sqlsrv_query($conn, $tsql, $params);
        if ($stmt) {
            echo "Tablas temporales vacías.<br>";
        } else {
            echo "Error al vaciar tablas temporales.<br>";
            die(print_r(sqlsrv_errors(), true));
        }



        /* Recorro cada proyecto en busca de la información del mes anterior */
        $tsql = "SELECT
            MIS_CAT_proyectos.CVE_CRE_IF,
            MIS_CAT_proyectos.CVE_CRE_ID_OFERTA,
            MIS_CAT_proyectos.NOM_PROYECTO,
            MIS_CAT_proyectos.NUM_REF_SHF,
            MIS_CAT_proyectos.NOM_PROMOTOR,
            MIS_CAT_proyectos.TIPO_CREDITO,
            MIS_CAT_proyectos.UBICACIÓN_EDO,
            MIS_CAT_proyectos.UBICACIÓN_MUN,
            MIS_CAT_proyectos.FECH_INI_CONTRATO,
            MIS_CAT_proyectos.LINEA_DE_CRE_POR_PROYECTO,
            MIS_CAT_proyectos.VALOR_PROYECTO,
            MIS_CAT_proyectos.TASA_INTERES,
            MIS_CAT_proyectos.VIV_TOTALES_PROYECTO,
            MIS_CAT_proyectos.COLATERAL,
            MIS_colaterales.NOM_PROYECTO,
            MIS_colaterales.FECH_COLATERAL,
            MIS_colaterales.FECH_FIN_CONTRATO,
            MIS_colaterales.AO_VIV_ACTIVAS,
            MIS_colaterales.VIV_LIB_CORTE_ANTERIOR,
            MIS_colaterales.VIV_LIB_PERIODO,
            MIS_colaterales.ACUM_VIV_LIB_FIN_P,
            MIS_colaterales.MONTO_MIN_ACUM_P_ANTERIOR,
            MIS_colaterales.MONTO_MIN_PERIODO,
            MIS_colaterales.MONTO_MIN_ACUM_FIN_P,
            MIS_colaterales.MONTO_POR_DISPONER,
            MIS_colaterales.MONTO_AMORT_ACUM_P_ANTERIOR,
            MIS_colaterales.MONTO_AMORTIZADO_PERIODO,
            MIS_colaterales.MONTO_AMORT_ACUM_FIN_P,
            MIS_colaterales.SALDO_INS_P_ANTERIOR,
            MIS_colaterales.SALDO_INS_CARTERA_FIN_P,
            MIS_colaterales.INT_COBRADOS_PERIODO,
            MIS_colaterales.COMISIONES_COBRADAS_PERIODO,
            MIS_colaterales.NUM_MESES_MOROSOS,
            MIS_colaterales.MONTO_INT_DEV_NO_CUBIERTOS
            FROM
            MIS_CAT_proyectos
            INNER JOIN MIS_colaterales
            ON
            MIS_CAT_proyectos.NOM_PROYECTO = MIS_colaterales.NOM_PROYECTO
            WHERE MIS_colaterales.FECH_COLATERAL = $mesPrevio";
        $stmt = sqlsrv_query( $conn, $tsql);
        if( $stmt === false)
        {
             echo "Error in query preparation/execution.\n";
             die( print_r( sqlsrv_errors(), true));
        }

        /* Retrieve each row as an associative array and display the results.*/
        while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC))
        {
            echo $row['FECH_COLATERAL'] . ", " . $row['MONTO_AMORT_ACUM_P_ANTERIOR'] . "<br>";
        }


        die("here");

        echo "<pre>";
        print_r($matriz);
        echo "</pre>";
        die("Terminado");

        $titulos = Array("CVE_CRED_IF", "CVE_CRED_ID_OFERTA", "NUM_REF_SHF", "NOM_CONJUNTO", "NOM_PROMOTOR", "TIPO_CREDITO", "UBICACION_ESTADO", "UBICACION_MUNICIPIO", "FECH_INI_CONTRATO", "LINEA_DE_CREDITO_POR_PROYECTO", "VALOR_PROYECTO", "TASA_INTERES", "VIVIENDAS_TOTALES_DEL_PROYECTO", "FECH_FIN_CONTRATO", "AO_VIV_ACTIVAS", "VIV_LIB_PERIODO", "MONTO_MIN_EN_EL_PERIODO", "MONTO_AMORT_EN_EL_PERIODO", "PROYECTO_MAY", "PROYECTO_MIN", "REP_MOR_INTERESES", "REP_INTPROV_INTERESES");
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($matriz, NULL, 'A2', true);
        $sheet->fromArray($titulos, NULL, 'A1', true);


        $letras =  Array("A","B","C","D","E","F","G","H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "V", "W");
        for ($i = 0; $i < sizeof($titulos); $i++)
        {
          $spreadsheet->getActiveSheet()->getColumnDimension($letras[$i])->setWidth(25);
        }

        $spreadsheet->getDefaultStyle()->getFont()->setName('Helvetica');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(13);



        $estiloColumnasEspecificas = [
                      'font' => [
                          'color' => array('rgb' => '000000'),
                          'size'  => 13,
                      ],
                      'alignment' => [
                          'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                      ],

                  ];
        $spreadsheet->getActiveSheet()->getStyle('A:V')->applyFromArray($estiloColumnasEspecificas);

        $styleArray = [
                      'font' => [
                          'bold' => true,
                          'color' => array('rgb' => 'FFFFFF'),
                          'size'  => 15,
                      ],
                      'alignment' => [
                          'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                      ],
                      'borders' => [
                          'outline' => [
                              'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                              'color' => array('argb' => 'FFFFFF'),
                          ],
                      ],
                      'fill' => [
                          'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                          'color' => [
                              'argb' => '2A7BD6',
                          ],
                      ],
                  ];

        $spreadsheet->getActiveSheet()->getStyle('A1:V1')->applyFromArray($styleArray);


        $writer = new Xlsx($spreadsheet);
        //$writer->save('helloworld.xlsx');

        if ($writer->save('reportes/colateral_' . date("d-m-Y_h.i.s_00") . '.xlsx'))
        {
            echo "<br>Error reporte";
            header('Location: autorizaAccesos.php?msg=reporteError');
        }
        else
        {
          $linkRep = 'reportes/colateral_' . date("d-m-Y_h.i.s_00") . '.xlsx';
          echo "<br>Reporte generado.";
          //echo "<script>window.location.href = 'reportesAccesos/REP_ACCESOS_' . $fechaActual . '.xlsx'';</script>";
          header('Location: reportes/colateral_' . date("d-m-Y_h.i.s_00") . '.xlsx');
        }

    } // END OF CARGA DE DATOS

?>
