const excel_file = document.getElementById('excel_file');
let sheet_data = [];

excel_file.addEventListener('change', (event) => {
  if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type)) {
    document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Only .xlsx or .xls file format is allowed</div>';
    excel_file.value = '';
    return false;
  }

  var reader = new FileReader();

  reader.readAsArrayBuffer(event.target.files[0]);

  reader.onload = function (event) {
    var data = new Uint8Array(reader.result);
    var work_book = XLSX.read(data, { type: 'array' });
    var sheet_name = work_book.SheetNames;
    sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], { header: 1 });

    if (sheet_data.length > 0) {
      renderTable();
    }

    excel_file.value = '';
  };
});

function renderTable() {
  var table_output = '<table class="table table-striped table-bordered">';

  for (var row = 0; row < sheet_data.length; row++) {
    table_output += '<tr>';

    for (var cell = 0; cell < sheet_data[row].length; cell++) {
      if (row === 0) {
        table_output += '<th>' + sheet_data[row][cell] + '</th>';
      } else {
        table_output += '<td><input type="text" value="' + sheet_data[row][cell] + '" onchange="updateData(' + row + ',' + cell + ', this.value)"></td>';
      }
    }

    table_output += '</tr>';
  }

  table_output += '</table>';

  // Add a "Download" button
  table_output += '<button onclick="downloadData()">Download</button>';

  document.getElementById('excel_data').innerHTML = table_output;
}

function updateData(row, cell, value) {
  sheet_data[row][cell] = value;
}

function downloadData() {
  var worksheet = XLSX.utils.aoa_to_sheet(sheet_data);
  var workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  var excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  saveAsExcelFile(excelBuffer, 'updated_data.xlsx');
}

function saveAsExcelFile(buffer, filename) {
  var blob = new Blob([buffer], { type: 'application/octet-stream' });
  var url = URL.createObjectURL(blob);
  var link = document.createElement('a');
  link.href = url;
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}


  
  

