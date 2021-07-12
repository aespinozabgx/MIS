// Call the dataTables jQuery plugin
$(document).ready(function() {
  $('#dataTable').DataTable({
    dom: 'Bfrtip',
      buttons: [
          {
              extend: 'copyHtml5',
              text: '<i class="far fa-copy"></i>',
              className: "btn btn-outline-primary btn-sm botonesHTML5"
          },
          {
              extend: 'excelHtml5',
              text: '2',
              className: "btn btn-outline-primary btn-sm botonesHTML5"
          },
          {
              extend: 'csvHtml5',
              text: '<i class="fas fa-file-csv"></i>',
              className: "btn btn-outline-primary btn-sm botonesHTML5"
          }
      ]
  });
});
