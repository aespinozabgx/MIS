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
        $mesActualziar = $_POST['toUpdate'];
        $mesActualziar = date("Y-m-d",strtotime($mesActualziar));

        if (validaCargaColateral($conn, $mesActualziar))
        {
          die("Ya existen datos cargados");
        }

        $fechaPorciones = explode(' ', $mesActualziar);
        $onlyDate = $fechaPorciones[0];
        $onlyDate = explode('-',$mesActualziar);
        $mes = $onlyDate[1];
        $mes -= 1;


        $monthNum = $mes;
        $dateObj = DateTime::createFromFormat('!m', $monthNum);
        $monthName = $dateObj->format('m');
        $preMonth = $onlyDate[0] . "-" . $monthName. "-01";

        echo "<br>Month to update: " .  $mesActualziar . "<br>";
        echo "<br>Prev month: " . $preMonth . "<br>";




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
        MIS_CAT_proyectos.LINEA_DE_CRE_POR_PROYECTO,
        MIS_CAT_proyectos.VALOR_PROYECTO,
        MIS_CAT_proyectos.TASA_INTERES,
        MIS_CAT_proyectos.VIV_TOTALES_PROYECTO,
        MIS_temp_shf.FECH_FIN_CONTRATO,
        MIS_temp_shf.AO_VIV_ACTIVAS,
        MIS_temp_shf.VIV_LIB_PERIODO,
        MIS_temp_shf.MONTO_MIN_EN_EL_PERIODO,
        MIS_temp_shf.MONTO_AMORT_EN_EL_PERIODO,

        /* monto amort acum fin periodo */
        /* acum viv lib a fin periodo */
        /* monto min acum fin periodo */
        /* monto por disponer */
        /* saldo ins cartera fin periodo */

        /* Viviendas Liberadas al Corte Anterior  */
        /* monto amort acum periodo ant */
        /* saldo ins periodo ant */
        /* monto min acum periodo ant */

        /* comisiones cobradas */
        MIS_temp_intprov.INTERESES AS INTERESES_COBRADOS_PERIODO,
        /* meses morosos */
        MIS_temp_morosidad.INT_DEVENGADO AS INTERESES_DEV_NO_CUBIERTOS


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


        $contador = 0;
        while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_BOTH))
        {

              $matriz[$contador]['COLATERAL']                      = $row['COLATERAL'];
              $matriz[$contador]['CVE_CRE_IF']                     = $row['CVE_CRE_IF'];
              $matriz[$contador]['CVE_CRE_ID_OFERTA']              = $row['CVE_CRE_ID_OFERTA'];
              $matriz[$contador]['NUM_REF_SHF']                    = $row['NUM_REF_SHF'];
              $matriz[$contador]['NOM_PROYECTO']                   = $row['NOM_PROYECTO'];
              $matriz[$contador]['NOM_PROMOTOR']                   = $row['NOM_PROMOTOR'];
              $matriz[$contador]['TIPO_CREDITO']                   = $row['TIPO_CREDITO'];
              $matriz[$contador]['UBICACIÓN_EDO']                  = $row['UBICACIÓN_EDO'];
              $matriz[$contador]['UBICACIÓN_MUN']                  = $row['UBICACIÓN_MUN'];
              $matriz[$contador]['FECH_INI_CONTRATO']              = $row['FECH_INI_CONTRATO'];
              $matriz[$contador]['LINEA_DE_CRE_POR_PROYECTO']      = $row['LINEA_DE_CRE_POR_PROYECTO'];
              $matriz[$contador]['VALOR_PROYECTO']                 = $row['VALOR_PROYECTO'];
              $matriz[$contador]['TASA_INTERES']                   = $row['TASA_INTERES'];
              $matriz[$contador]['VIV_TOTALES_PROYECTO']           = $row['VIV_TOTALES_PROYECTO'];
              $matriz[$contador]['FECH_FIN_CONTRATO']              = $row['FECH_FIN_CONTRATO'];
              $matriz[$contador]['AO_VIV_ACTIVAS']                 = $row['AO_VIV_ACTIVAS'];
              $matriz[$contador]['VIV_LIB_PERIODO']                = $row['VIV_LIB_PERIODO'];

              /* shf */
              $matriz[$contador]['MONTO_MIN_EN_EL_PERIODO']        = $row['MONTO_MIN_EN_EL_PERIODO'];
              $matriz[$contador]['MONTO_AMORT_EN_EL_PERIODO']      = $row['MONTO_AMORT_EN_EL_PERIODO'];
              $matriz[$contador]['INTERESES_COBRADOS_PERIODO']     = $row['INTERESES_COBRADOS_PERIODO'];
              $matriz[$contador]['INTERESES_DEV_NO_CUBIERTOS']     = $row['INTERESES_DEV_NO_CUBIERTOS'];

              $contador++;

        }

        /* Elimino tablas temporales */
        $tsql = "DELETE FROM MIS_temp_intprov; DELETE FROM MIS_temp_morosidad; DELETE FROM MIS_temp_shf";

        /* Set parameter values. */
        $params = array();

        /* Prepare and execute the query. */
        $stmt = sqlsrv_query($conn, $tsql, $params);
        if ($stmt)
        {
            echo "Tablas temporales vacías.<br>";
        }
        else
        {
            echo "Error al vaciar tablas temporales.<br>";
            die(print_r(sqlsrv_errors(), true));
        }


        /*UNO LA INFORAMCIÓN */

        $tsql = "EXEC MIS_obtieneMesPrevio ?";
        $params = Array($preMonth);
        $stmt = sqlsrv_query($conn, $tsql, $params);
        if( $stmt === false)
        {
             echo "Error in query execution.<br>";
             die( print_r( sqlsrv_errors(), true));
        }

        /* Retrieve each row as an associative array and display the results.*/
        $c = 0;
        $datosMesPrevio = Array();
        while( $row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC))
        {
            echo $row['NOM_PROYECTO'] . "<br>";
            $datosMesPrevio[$c]['NOM_PROYECTO']                   = $row['NOM_PROYECTO'];
            $datosMesPrevio[$c]['FECH_COLATERAL']                 = date_format($row['FECH_COLATERAL'], "d-m-Y");
            $datosMesPrevio[$c]['COLATERAL']                      = $row['COLATERAL'];
            $datosMesPrevio[$c]['VIV_LIB_CORTE_ANTERIOR']         = $row['VIV_LIB_CORTE_ANTERIOR'];
            $datosMesPrevio[$c]['ACUM_VIV_LIB_FIN_P']             = $row['ACUM_VIV_LIB_FIN_P'];
            $datosMesPrevio[$c]['MONTO_MIN_ACUM_P_ANTERIOR']      = $row['MONTO_MIN_ACUM_P_ANTERIOR'];
            $datosMesPrevio[$c]['MONTO_MIN_ACUM_FIN_P']           = $row['MONTO_MIN_ACUM_FIN_P'];
            $datosMesPrevio[$c]['MONTO_POR_DISPONER']             = $row['MONTO_POR_DISPONER'];
            $datosMesPrevio[$c]['MONTO_AMORT_ACUM_P_ANTERIOR']    = $row['MONTO_AMORT_ACUM_P_ANTERIOR'];
            $datosMesPrevio[$c]['MONTO_AMORT_ACUM_FIN_P']         = $row['MONTO_AMORT_ACUM_FIN_P'];
            $datosMesPrevio[$c]['SALDO_INS_P_ANTERIOR']           = $row['SALDO_INS_P_ANTERIOR'];
            $datosMesPrevio[$c]['SALDO_INS_CARTERA_FIN_P']        = $row['SALDO_INS_CARTERA_FIN_P'];
            $datosMesPrevio[$c]['COMISIONES_COBRADAS_PERIODO']    = $row['COMISIONES_COBRADAS_PERIODO'];
            $datosMesPrevio[$c]['NUM_MESES_MOROSOS']              = $row['NUM_MESES_MOROSOS'];
            $c++;
        }

        /* FIN UNO LA INFORMACIÓN */
        $matriz_row = count($matriz);
        $matriz_col = max(array_map('count', $matriz));

        $datosMesPrevio_row = count($matriz);
        $datosMesPrevio_col = max(array_map('count', $matriz));


        for ($cont=0; $cont < $matriz_row; $cont++)
        {
            //echo "Buscando " . $matriz[$cont][4] . "<br>";
            for ($i=0; $i < $datosMesPrevio_row; $i++)
            {
                if ($datosMesPrevio[$i]['NOM_PROYECTO'] == $matriz[$cont]['NOM_PROYECTO'])
                {
                  echo "<hr><br>";
                    echo $cont . ". Encontrado: " . $datosMesPrevio[$i]['NOM_PROYECTO'] . " - " . $matriz[$cont]['NOM_PROYECTO'] . "<br><br>";

                    echo $mesActualziar;
                    echo $matriz[$cont]['COLATERAL'];
                    echo $matriz[$cont]['CVE_CRE_IF'];
                    echo $matriz[$cont]['CVE_CRE_ID_OFERTA'];
                    echo $matriz[$cont]['NUM_REF_SHF'];
                    echo $matriz[$cont]['NOM_PROYECTO'];
                    echo $matriz[$cont]['NOM_PROMOTOR'];
                    echo $matriz[$cont]['TIPO_CREDITO'];
                    echo $matriz[$cont]['UBICACIÓN_EDO'];
                    echo $matriz[$cont]['UBICACIÓN_MUN'];
                    echo date_format($matriz[$cont]['FECH_INI_CONTRATO'], "d-m-Y");
                    echo $matriz[$cont]['LINEA_DE_CRE_POR_PROYECTO'];
                    echo $matriz[$cont]['VALOR_PROYECTO'];
                    echo $matriz[$cont]['TASA_INTERES'];
                    echo $matriz[$cont]['VIV_TOTALES_PROYECTO'];
                    echo date_format($matriz[$cont]['FECH_FIN_CONTRATO'], "d-m-Y");
                    echo $matriz[$cont]['AO_VIV_ACTIVAS'];
                    echo $matriz[$cont]['VIV_LIB_PERIODO'];
                    echo $matriz[$cont]['MONTO_MIN_EN_EL_PERIODO'];
                    echo $matriz[$cont]['MONTO_AMORT_EN_EL_PERIODO'];

                    /* SUMA DE COLUMNAS */
                    echo ($datosMesPrevio[$i]['MONTO_AMORT_ACUM_P_ANTERIOR'] + $matriz[$cont]['MONTO_AMORT_EN_EL_PERIODO']);
                    echo ($datosMesPrevio[$i]['VIV_LIB_CORTE_ANTERIOR']      + $matriz[$cont]['VIV_LIB_PERIODO']);
                    echo ($datosMesPrevio[$i]['MONTO_MIN_ACUM_P_ANTERIOR']   + $matriz[$cont]['MONTO_MIN_EN_EL_PERIODO']);
                    echo ($datosMesPrevio[$i]['MONTO_MIN_ACUM_FIN_P']        + $matriz[$cont]['MONTO_MIN_EN_EL_PERIODO']);
                    echo (($matriz[$cont]['MONTO_MIN_EN_EL_PERIODO']         - $matriz[$cont]['MONTO_AMORT_EN_EL_PERIODO']) + $datosMesPrevio[$i]['SALDO_INS_P_ANTERIOR']);
                    /* FIN SUMA DE COLUMNAS */

                    echo $datosMesPrevio[$i]['VIV_LIB_CORTE_ANTERIOR'];
                    echo $datosMesPrevio[$i]['MONTO_AMORT_ACUM_FIN_P'];
                    echo $datosMesPrevio[$i]['SALDO_INS_CARTERA_FIN_P'];
                    echo $datosMesPrevio[$i]['MONTO_MIN_ACUM_P_ANTERIOR'];
                    echo $matriz[$cont]['INTERESES_COBRADOS_PERIODO'];
                    echo $datosMesPrevio[$i]['NUM_MESES_MOROSOS'];
                    echo $matriz[$cont]['INTERESES_DEV_NO_CUBIERTOS'];

                    // echo "<br><b>NOM_PROYECTO:</b> " . $datosMesPrevio[$i]['NOM_PROYECTO'];
                    // echo "<br><b>FECH_COLATERAL:</b> " . "Fecha: " . $datosMesPrevio[$i]['FECH_COLATERAL'];
                    // echo "<br><b>MONTO_POR_DISPONER:</b> " . $datosMesPrevio[$i]['MONTO_POR_DISPONER'];
                    // echo "<br><b>ACUM_VIV_LIB_FIN_P:</b> " . $datosMesPrevio[$i]['ACUM_VIV_LIB_FIN_P'];
                    // echo "<br><b>COMISIONES_COBRADAS_PERIODO:</b> " . $datosMesPrevio[$i]['COMISIONES_COBRADAS_PERIODO'];

                    echo " <br><br> ";

                }
            }
        }

        die("<br><br>Terminado");

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
