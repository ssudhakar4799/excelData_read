document.getElementById('uploadForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent default form submission

    var fileInput = document.getElementById('excelFile');
    var file = fileInput.files[0];

    if (!file) {
        alert('Please select an Excel file.');
        return;
    }

    var reader = new FileReader();
    reader.onload = function(e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });
        var sheetName = workbook.SheetNames[0];
        var sheet = workbook.Sheets[sheetName];
        var excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        console.log(excelData);
        displayExcelData(excelData);
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('addDataForm').addEventListener('submit', function(event) {
    event.preventDefault(); // Prevent default form submission

    var name = document.getElementById('name').value;
    var email = document.getElementById('email').value;
    var mobile = document.getElementById('mobile').value;

    var newRow = [name, email, mobile];
    addDataToExcel(newRow);
});

function displayExcelData(data) {
    var table = document.getElementById('excelDataTable');
    table.style.width = "100%";
    table.innerHTML = ''; // Clear previous data
    data.forEach(function(row) {
        var newRow = table.insertRow();
        newRow.style.border = "1px solid red";
        newRow.style.color = "red";
        newRow.style.fontSize = "larger";
        newRow.style.fontFamily = "cursive";
        row.forEach(function(cellData) {
            var newCell = newRow.insertCell();
            newCell.appendChild(document.createTextNode(cellData));
        });
    });
}

function addDataToExcel(newRow) {
    var table = document.getElementById('excelDataTable');
    var newRowElem = table.insertRow();
    newRowElem.style.border = "1px solid red";
    newRowElem.style.color = "red";
    newRowElem.style.fontSize = "larger";
    newRowElem.style.fontFamily = "cursive";
    
    newRow.forEach(function(cellData) {
        var newCell = newRowElem.insertCell();
        newCell.appendChild(document.createTextNode(cellData));
    });
}
