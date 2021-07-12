<?php

  require 'php/conexion.php';
  require 'php/funciones.php';

  date_default_timezone_set('America/Mexico_City');
  setlocale(LC_ALL,'es_ES');




?>
<!DOCTYPE html>
<html lang="es">

<head>

    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>CONCRECES</title>

    <!-- Custom fonts for this template -->
    <link href="vendor/fontawesome-free/css/all.min.css" rel="stylesheet" type="text/css">
    <link
        href="https://fonts.googleapis.com/css?family=Nunito:200,200i,300,300i,400,400i,600,600i,700,700i,800,800i,900,900i"
        rel="stylesheet">

    <link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.css" />

    <!-- Custom styles for this template -->
    <link href="css/sb-admin-2.min.css" rel="stylesheet">

    <!-- Custom styles for this page -->
    <link href="vendor/datatables/dataTables.bootstrap4.min.css" rel="stylesheet">

    <style type="text/css">

      td
      {
        font-size: 12px;
      }

      .botonesHTML5
      {
        margin: 0px;
        position:relative;
      }

    </style>

    </style>
</head>

<body id="page-top">

    <!-- Page Wrapper -->
    <div id="wrapper">

        <?php
          if (file_exists("navbar.php"))
          {
            //require 'navbar.php';
          }
        ?>

        <!-- Content Wrapper -->
        <div id="content-wrapper" class="d-flex flex-column">

            <!-- Main Content -->
            <div id="content">

                <!-- Topbar -->
                <nav class="navbar navbar-expand navbar-light bg-white topbar mb-4 static-top shadow">

                    <!-- Topbar Navbar -->
                    <ul class="navbar-nav ml-auto">


                        <!-- Nav Item - Messages -->


                        <div class="topbar-divider d-none d-sm-block"></div>

                        <!-- Nav Item - User Information -->
                        <li class="nav-item dropdown no-arrow">
                            <a class="nav-link dropdown-toggle" href="#" id="userDropdown" role="button"
                                data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                <span class="mr-2 d-none d-lg-inline text-gray-600 small">Usuario</span>
                                <img class="img-profile rounded-circle"
                                    src="img/undraw_profile.svg">
                            </a>
                            <!-- Dropdown - User Information -->
                            <div class="dropdown-menu dropdown-menu-right shadow animated--grow-in"
                                aria-labelledby="userDropdown">
                                <a class="dropdown-item" href="#">
                                    <i class="fas fa-user fa-sm fa-fw mr-2 text-gray-400"></i>
                                    Profile
                                </a>
                                <a class="dropdown-item" href="#">
                                    <i class="fas fa-cogs fa-sm fa-fw mr-2 text-gray-400"></i>
                                    Settings
                                </a>
                                <a class="dropdown-item" href="#">
                                    <i class="fas fa-list fa-sm fa-fw mr-2 text-gray-400"></i>
                                    Activity Log
                                </a>
                                <div class="dropdown-divider"></div>
                                <a class="dropdown-item" href="#" data-toggle="modal" data-target="#logoutModal">
                                    <i class="fas fa-sign-out-alt fa-sm fa-fw mr-2 text-gray-400"></i>
                                    Logout
                                </a>
                            </div>
                        </li>

                    </ul>

                </nav>
                <!-- End of Topbar -->

                <!-- Begin Page Content -->
                <div class="container-fluid">

                    <!-- Page Heading -->
                    <h1 class="h3 mb-2 text-gray-800">Colaterales</h1>
                    <p class="mb-4"></p>
                    <!-- DataTales Example -->
                    <div class="card shadow mb-4" id="tablaBusqueda">

                        <div class="card-header py-3 d-flex flex-row align-items-center justify-content-between">

                          <h6 class="m-0 font-weight-bold text-primary">PROYECTOS </h6><div id="output" class="m-0 font-weight-bold text-danger"></div>

                          <div class=" " role="group">

                              <button type="button" class="btn btn-white text-primary btn-circle" data-toggle="modal" data-target="#modalFiltros">
                                  <i class="fas fa-filter"></i>
                              </button>

                              <button type="button" class="btn btn-white text-primary btn-circle" data-toggle="modal" data-target="#modalCarga">
                                  <i class="fas fa-cloud-upload-alt"></i>
                              </button>

                          </div>

                        </div>


                        <div class="card-body">
                            <div class="table-responsive">

                              <?php


                                $tsql = "   ";


                                if (isset($_GET['date']))
                                {
                                    $tsql .= " WHERE MONTH(FECH_COLATERAL) = DATEPART(m,'" . $_GET['date'] . "')";
                                }

                                //echo "<br>sql: " .$tsql;
                                ?>

                                <table class="table table-bordered" id="dataTable" width="100%" cellspacing="0"  data-order='[[ 1, "desc" ]]' data-page-length='25'>
                                    <thead>

                                        <tr class="text-primary text-decoration-none" style="font-size: 10px;">
                                          <th>NUM COLATERAL</th>
                                          <th>FECH_COLATERAL</th>
                                          <th>CVE_CRE_IF</th>
                                          <th>CVE_CRE_ID_OFERTA</th>
                                          <th>NUM_REF_SHF</th>
                                          <th>NOM_PROYECTO</th>
                                          <th>NOM_PROMOTOR</th>
                                          <th>TIPO_CREDITO</th>
                                          <th>UBICACIÓN_EDO</th>
                                          <th>UBICACIÓN_MUN</th>
                                          <th>FECH_INI_CONTRATO</th>
                                          <th>LINEA_DE_CRE_POR_PROYECTO</th>
                                          <th>VALOR_PROYECTO</th>
                                          <th>TASA_INTERES</th>
                                          <th>VIV_TOTALES_PROYECTO</th>
                                          <th>FECH_FIN_CONTRATO</th>
                                          <th>AO_VIV_ACTIVAS</th>
                                          <th>VIV_LIB_PERIODO</th>
                                          <th>MONTO_MIN_PERIODO</th>
                                          <th>MONTO_AMORTIZADO_PERIODO</th>
                                          <th>MONTO_AMORT_ACUM_FIN_P</th>
                                          <th>ACUM_VIV_LIB_FIN_P</th>
                                          <th>MONTO_MIN_ACUM_FIN_P</th>
                                          <th>MONTO_POR_DISPONER</th>
                                          <th>SALDO_INS_CARTERA_FIN_P</th>
                                          <th>VIV_LIB_CORTE_ANTERIOR</th>
                                          <th>MONTO_AMORT_ACUM_P_ANTERIOR</th>
                                          <th>SALDO_INS_P_ANTERIOR</th>
                                          <th>MONTO_MIN_ACUM_P_ANTERIOR</th>
                                          <th>INT_COBRADOS_PERIODO</th>
                                          <th>NUM_MESES_MOROSOS</th>
                                          <th>MONTO_INT_DEV_NO_CUBIERTOS</th>
                                        </tr>

                                    </thead>
                                    <tfoot>
                                      <tr style="font-size: 10px;">
                                        <th>NUM COLATERAL</th>
                                        <th>FECH_COLATERAL</th>
                                        <th>CVE_CRE_IF</th>
                                        <th>CVE_CRE_ID_OFERTA</th>
                                        <th>NUM_REF_SHF</th>
                                        <th>NOM_PROYECTO</th>
                                        <th>NOM_PROMOTOR</th>
                                        <th>TIPO_CREDITO</th>
                                        <th>UBICACIÓN_EDO</th>
                                        <th>UBICACIÓN_MUN</th>
                                        <th>FECH_INI_CONTRATO</th>
                                        <th>LINEA_DE_CRE_POR_PROYECTO</th>
                                        <th>VALOR_PROYECTO</th>
                                        <th>TASA_INTERES</th>
                                        <th>VIV_TOTALES_PROYECTO</th>
                                        <th>FECH_FIN_CONTRATO</th>
                                        <th>AO_VIV_ACTIVAS</th>
                                        <th>VIV_LIB_PERIODO</th>
                                        <th>MONTO_MIN_PERIODO</th>
                                        <th>MONTO_AMORTIZADO_PERIODO</th>
                                        <th>MONTO_AMORT_ACUM_FIN_P</th>
                                        <th>ACUM_VIV_LIB_FIN_P</th>
                                        <th>MONTO_MIN_ACUM_FIN_P</th>
                                        <th>MONTO_POR_DISPONER</th>
                                        <th>SALDO_INS_CARTERA_FIN_P</th>
                                        <th>VIV_LIB_CORTE_ANTERIOR</th>
                                        <th>MONTO_AMORT_ACUM_P_ANTERIOR</th>
                                        <th>SALDO_INS_P_ANTERIOR</th>
                                        <th>MONTO_MIN_ACUM_P_ANTERIOR</th>
                                        <th>INT_COBRADOS_PERIODO</th>
                                        <th>NUM_MESES_MOROSOS</th>
                                        <th>MONTO_INT_DEV_NO_CUBIERTOS</th>
                                      </tr>
                                    </tfoot>

                                    <tbody>
                                        <?php


                                          $stmt = sqlsrv_query( $conn, $tsql );
                                          if( $stmt === false )
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
                                          else
                                          {
                                            $conteo = 0;
                                            while( $row = sqlsrv_fetch_array( $stmt, SQLSRV_FETCH_ASSOC))
                                            {
                                                $conteo++;
                                                echo "<tr>";
                                                echo "<td>" . $row['COLATERAL'] . "</td>";
                                                echo "<td>" . (date_format($row['FECH_COLATERAL'], "m-Y")) . "</td>";
                                                echo "<td>" . $row['CVE_CRE_IF'] . "</td>";
                                                echo "<td>" . $row['CVE_CRE_ID_OFERTA'] . "</td>";
                                                echo "<td>" . $row['NUM_REF_SHF'] . "</td>";
                                                echo "<td>" . $row['NOM_PROYECTO'] . "</td>";
                                                echo "<td>" . $row['NOM_PROMOTOR'] . "</td>";
                                                echo "<td>" . $row['TIPO_CREDITO'] . "</td>";
                                                echo "<td>" . $row['UBICACIÓN_EDO'] . "</td>";
                                                echo "<td>" . $row['UBICACIÓN_MUN'] . "</td>";
                                                echo "<td>" . date_format($row['FECH_INI_CONTRATO'], "d-m-Y") . "</td>";
                                                echo "<td>" . $row['LINEA_DE_CRE_POR_PROYECTO'] . "</td>";
                                                echo "<td>" . $row['VALOR_PROYECTO'] . "</td>";
                                                echo "<td>" . $row['TASA_INTERES'] . "</td>";
                                                echo "<td>" . $row['VIV_TOTALES_PROYECTO'] . "</td>";
                                                echo "<td>" . date_format($row['FECH_FIN_CONTRATO'], "d-m-Y") . "</td>";
                                                echo "<td>" . $row['AO_VIV_ACTIVAS'] . "</td>";
                                                echo "<td>" . $row['VIV_LIB_PERIODO'] . "</td>";
                                                echo "<td>" . $row['MONTO_MIN_PERIODO'] . "</td>";
                                                echo "<td>" . $row['MONTO_AMORTIZADO_PERIODO'] . "</td>";
                                                echo "<td>" . $row['MONTO_AMORT_ACUM_FIN_P'] . "</td>";
                                                echo "<td>" . $row['ACUM_VIV_LIB_FIN_P'] . "</td>";
                                                echo "<td>" . $row['MONTO_MIN_ACUM_FIN_P'] . "</td>";
                                                echo "<td>" . $row['MONTO_POR_DISPONER'] . "</td>";
                                                echo "<td>" . $row['SALDO_INS_CARTERA_FIN_P'] . "</td>";
                                                echo "<td>" . $row['VIV_LIB_CORTE_ANTERIOR'] . "</td>";
                                                echo "<td>" . $row['MONTO_AMORT_ACUM_P_ANTERIOR'] . "</td>";
                                                echo "<td>" . $row['SALDO_INS_P_ANTERIOR'] . "</td>";
                                                echo "<td>" . $row['MONTO_MIN_ACUM_P_ANTERIOR'] . "</td>";
                                                echo "<td>" . $row['INT_COBRADOS_PERIODO'] . "</td>";
                                                echo "<td>" . $row['NUM_MESES_MOROSOS'] . "</td>";
                                                echo "<td>" . $row['MONTO_INT_DEV_NO_CUBIERTOS'] . "</td>";
                                                echo "</tr>";

                                            }

                                            /* Free statement and connection resources. */
                                            sqlsrv_free_stmt( $stmt);
                                            sqlsrv_close( $conn);

                                          }

                                          if (isset($conteo) && $conteo == 0 && isset($_GET['date']) && strtotime($_GET['date']))
                                          {
                                            ?>

                                              <script>
                                                var txt = '<?php echo date('m-Y', strtotime($_GET['date'])); ?>';
                                                document.getElementById('output').innerHTML = "No se encontró información de "  + txt;
                                              </script>

                                            <?php
                                          }
                                        ?>


                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                </div>
                <!-- /.container-fluid -->

            </div>
            <!-- End of Main Content -->

            <!-- Footer -->
            <footer class="sticky-footer bg-white">
                <div class="container my-auto">
                    <div class="copyright text-center my-auto">
                        <span>Copyright &copy; Your Website 2020</span>
                    </div>
                </div>
            </footer>
            <!-- End of Footer -->

        </div>
        <!-- End of Content Wrapper -->

    </div>
    <!-- End of Page Wrapper -->

    <!-- Scroll to Top Button-->
    <a class="scroll-to-top rounded" href="#page-top">
        <i class="fas fa-angle-up"></i>
    </a>

    <!-- Logout Modal-->
    <div class="modal fade" id="logoutModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel"
        aria-hidden="true">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalLabel">Ready to Leave?</h5>
                    <button class="close" type="button" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">×</span>
                    </button>
                </div>
                <div class="modal-body">Select "Logout" below if you are ready to end your current session.</div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" type="button" data-dismiss="modal">Cancel</button>
                    <a class="btn btn-primary" href="login.html">Logout</a>
                </div>
            </div>
        </div>
    </div>

    <?php
      if (file_exists("modals.php"))
      {
        require "modals.php";
      }
    ?>
    <!-- Bootstrap core JavaScript-->
    <script src="vendor/jquery/jquery.min.js"></script>
    <script src="vendor/bootstrap/js/bootstrap.bundle.min.js"></script>

    <!-- Core plugin JavaScript-->
    <script src="vendor/jquery-easing/jquery.easing.min.js"></script>

    <!-- Custom scripts for all pages-->
    <script src="js/sb-admin-2.min.js"></script>

    <!-- Page level plugins -->
    <script src="vendor/datatables/jquery.dataTables.min.js"></script>
    <script src="vendor/datatables/dataTables.bootstrap4.min.js"></script>
    <script src="vendor/datatables/dataTables.buttons.js"></script>
    <script src="vendor/datatables/buttons.html5.min.js"></script>


    <!-- Page level custom scripts -->
    <script src="js/demo/datatables-demo.js"></script>

    <!--  daterangePicker -->
    <script type="text/javascript" src="https://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.min.js"></script>

    <script>

      $(function() {

        var d = new Date();
        var month = d.getMonth()+1;
        var day = d.getDate();
        var output = d.getFullYear() + '-' +
            ((''+month).length<2 ? '0' : '') + month;


        $('input[name="daterange"]').daterangepicker({
          singleDatePicker: true,
          startDate: new Date(),
          locale: {
            format: 'MM-YYYY',
            applyLabel: "Aplicar",
            cancelLabel: "Cancelar",
            daysOfWeek: [
                "Do",
                "Lu",
                "Ma",
                "Mi",
                "Ju",
                "Vi",
                "Sa"
            ],
            monthNames: [
                "Enero",
                "Febrero",
                "Marzo",
                "Abril",
                "Mayo",
                "Junio",
                "Julio",
                "Agosto",
                "Septiembre",
                "Octubre",
                "Noviembre",
                "Diciembre"
            ],
          }
        });
      });


    </script>

</body>

</html>
