document.addEventListener('DOMContentLoaded', function () {
    const inputElement = document.getElementById('input-excel');
    if (inputElement) {
        // Cho phép upload nhiều file
        inputElement.addEventListener('change', handleFiles, false);
    } else {
        console.error('Không tìm thấy phần tử với ID "input-excel"');
    }
});

let allData = {}; // Biến để lưu trữ dữ liệu của tất cả các file và sheet

function handleFiles(e) {
    const files = e.target.files;
    if (!files.length) {
        alert('Vui lòng chọn ít nhất một tệp Excel.');
        return;
    }

    const excelDataDiv = document.getElementById('excel-data');
    excelDataDiv.innerHTML = ''; // Xóa nội dung cũ trước khi bắt đầu

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const reader = new FileReader();

        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Lọc các sheet không ẩn
            const visibleSheets = workbook.SheetNames.filter(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                return !(sheet['!hidden'] === 1); // Sheet ẩn có thuộc tính '!hidden' === 1
            });

            visibleSheets.forEach(function (sheetName) {
                const sheetContainer = document.createElement('div');
                sheetContainer.classList.add('sheet-container');

                const sheetHeader = document.createElement('div');
                sheetHeader.classList.add('sheet-header');

                const sheetTitle = document.createElement('h2');
                sheetTitle.textContent = 'Sheet: ' + sheetName;

                const toggleButton = document.createElement('button');
                toggleButton.classList.add('toggle-button');
                toggleButton.innerHTML = '<i class="fas fa-plus"></i>';

                toggleButton.addEventListener('click', function () {
                    if (sheetContent.style.display === 'none') {
                        sheetContent.style.display = 'block';
                        toggleButton.innerHTML = '<i class="fas fa-minus"></i>';
                    } else {
                        sheetContent.style.display = 'none';
                        toggleButton.innerHTML = '<i class="fas fa-plus"></i>';
                    }
                });

                sheetHeader.appendChild(sheetTitle);
                sheetHeader.appendChild(toggleButton);
                sheetContainer.appendChild(sheetHeader);

                const sheetContent = document.createElement('div');
                sheetContent.classList.add('sheet-content');
                sheetContent.style.display = 'none';

                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (jsonData.length === 0) {
                    const noDataMsg = document.createElement('p');
                    noDataMsg.textContent = 'Sheet này không có dữ liệu.';
                    sheetContent.appendChild(noDataMsg);
                } else {
                    const table = document.createElement('table');
                    table.classList.add('excel-table');
                    const tbody = document.createElement('tbody');

                    if (jsonData[0]) {
                        const headerRow = document.createElement('tr');
                        jsonData[0].forEach(function (cellData) {
                            const th = document.createElement('th');
                            th.textContent = cellData !== undefined ? cellData : '';
                            headerRow.appendChild(th);
                        });
                        table.appendChild(headerRow);
                    }

                    for (let i = 1; i < jsonData.length; i++) {
                        const rowData = jsonData[i];
                        const row = document.createElement('tr');
                        rowData.forEach(function (cellData) {
                            const td = document.createElement('td');
                            td.textContent = cellData !== undefined ? cellData : '';
                            row.appendChild(td);
                        });
                        tbody.appendChild(row);
                    }

                    table.appendChild(tbody);
                    sheetContent.appendChild(table);
                }

                sheetContainer.appendChild(sheetContent);
                excelDataDiv.appendChild(sheetContainer);

                // Lưu dữ liệu vào bảng
                if (!allData[file.name]) {
                    allData[file.name] = {};
                }
                allData[file.name][sheetName] = jsonData;

                console.log(`Data saved for file: ${file.name}, sheet: ${sheetName}`);
            });
        };

        reader.onerror = function (ex) {
            console.error('Error reading file', ex);
            alert('Đã xảy ra lỗi khi đọc tệp.');
        };

        // Đọc file dưới dạng array buffer thay vì binary string
        reader.readAsArrayBuffer(file);
    }
}

// Hàm để lấy dữ liệu cho design report
function getDataForReport() {
    return allData;
}
