
<!-- Modal -->
<div class="modal fade" id="modalFiltros" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">Filtro</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>

      <form class="" action="procesaMIS.php" method="post">
        <div class="modal-body">
          <p>
            <label>Fecha al corte:</label>
            <input type="date" class="form-control" name="fechaReporte" value="">
          </p>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancelar</button>
          <button type="submit" class="btn btn-primary" name="btnFiltra">Aplicar filtro</button>
        </div>
      </form>

    </div>
  </div>
</div>



<!-- modal Carga -->
<div class="modal fade" id="modalCarga" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <form class="" action="procesaMIS.php" method="post" enctype="multipart/form-data">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title" id="exampleModalLongTitle">Carga de información</h5>
          <button type="button" class="close" data-dismiss="modal" aria-label="Close">
            <span aria-hidden="true">&times;</span>
          </button>
        </div>
        <div class="modal-body">

            <div class="row">
              <div class="col-sm">

                <label class="text-primary">Mes a actualizar:</label>

                <div class="input-group mb-3">
                  <div class="input-group-prepend">
                    <span class="input-group-text" id="basic-addon1"><i class="far fa-calendar-alt"></i></span>
                  </div>
                  <input type="date" class="form-control" pattern="" name="toUpdate" value="<?php echo date('Y-m-d'); ?>" required>
                </div>

              </div>
            </div>


            <div class="row">
              <div class="col-sm">

                <div class="input-group mb-3">
                  <label for="RepPagoInt" class="text-primary">Reporte pago de intereses y provisiones</label>
                  <input type="file" class="form-control-file" name="repIntProv" id="RepPagoInt" required>
                </div>

              </div>
            </div>

            <div class="row">
              <div class="col-sm">

                  <label for="RepMorosidad" class="text-primary">Reporte de morosidad</label>
                  <input type="file" class="form-control-file" name="reporteMorosidad" required>

              </div>
            </div>

            <div class="row">
              <div class="col-sm">

                  <label for="RepMorosidad" class="text-primary">Reporte de SHF</label>
                  <input type="file" class="form-control-file" name="reporteSHF" required>

              </div>
            </div>

        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancelar</button>
          <button type="submit" class="btn btn-primary" name="cargaDatos">Actualizar</button>
        </div>
      </div>
    </form>
  </div>
</div>
