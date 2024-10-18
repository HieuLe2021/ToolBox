let workbooks = {};  // Lưu workbook của tất cả các file
let fileIDs = {};  // Lưu ID của các file

document.addEventListener('DOMContentLoaded', function () {
    const inputElement = document.getElementById('input-excel');
    if (inputElement) {
        inputElement.addEventListener('change', handleFile, false);
    } else {
        console.error('Không tìm thấy phần tử với ID "input-excel"');
    }
});

// Hàm import file
function handleFile(e) {
    const files = e.target.files;
    if (!files.length) {
        alert('Vui lòng chọn ít nhất một tệp Excel.');
        return;
    }

    const excelDataDiv = document.getElementById('excel-data');
    excelDataDiv.innerHTML = '';  // Xóa nội dung cũ

    Array.from(files).forEach((file, index) => {
        const reader = new FileReader();
        const fileName = file.name.replace('.xlsx', '');  // Loại bỏ phần mở rộng .xlsx
        const fileID = `file_${index + 1}`;  // Tạo ID cho file

        reader.onload = function (e) {
            const data = e.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });

            // Lấy dữ liệu hiện tại từ localStorage (nếu có)
            let existingData = JSON.parse(localStorage.getItem(fileID)) || {};

            // Duyệt qua từng sheet trong workbook
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName === 'hiddenSheet') return;  // Bỏ qua hiddenSheet nếu có

                const worksheet = workbook.Sheets[sheetName];
                let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Format lại dữ liệu JSON thành array of objects
                const columns = jsonData[0];  // Lấy tên các cột
                const formattedRows = jsonData.slice(1).map(row => {
                    const rowObject = {};
                    row.forEach((cellData, index) => {
                        rowObject[columns[index]] = cellData;
                    });
                    return rowObject;
                });

                // Lưu sheet vào `existingData`, thêm mới hoặc cập nhật nếu đã có
                existingData[sheetName] = formattedRows;

                // Tiếp tục phần hiển thị dữ liệu lên giao diện
                const sheetContainer = document.createElement('div');
                sheetContainer.classList.add('sheet-container');
                sheetContainer.setAttribute('data-file-id', fileID);  // Lưu ID file vào div

                const sheetHeader = document.createElement('div');
                sheetHeader.classList.add('sheet-header');

                const sheetTitle = document.createElement('h2');
                sheetTitle.textContent = `${fileName}_${sheetName}`;  // Tên file sẽ là namefile_sheet

                const toggleButton = document.createElement('button');
                toggleButton.classList.add('toggle-button');
                toggleButton.innerHTML = '<i class="fas fa-plus"></i>';

                const sheetContent = document.createElement('div');
                sheetContent.classList.add('sheet-content');
                sheetContent.style.display = 'none';

                toggleButton.addEventListener('click', function () {
                    if (sheetContent.style.display === 'none') {
                        sheetContent.style.display = 'block';
                        toggleButton.innerHTML = '<i class="fas fa-minus"></i>';
                    } else {
                        sheetContent.style.display = 'none';
                        toggleButton.innerHTML = '<i class="fas fa-plus"></i>';
                    }
                });

                const addButton = document.createElement('button');
                addButton.classList.add('add-button');
                addButton.textContent = 'Thêm Dữ Liệu';

                // Thêm sự kiện cho nút "Thêm Dữ Liệu"
                addButton.addEventListener('click', function () {
                    showAddDataPopup(fileID, sheetName, jsonData[1]);  // Pass file ID và sheet name cho popup
                });

                const buttonContainer = document.createElement('div');
                buttonContainer.classList.add('button-container');
                buttonContainer.appendChild(toggleButton);
                buttonContainer.appendChild(addButton);

                sheetHeader.appendChild(sheetTitle);
                sheetHeader.appendChild(buttonContainer);
                sheetContainer.appendChild(sheetHeader);

                // Kiểm tra nếu không có dữ liệu
                if (jsonData.length === 0) {
                    const noDataMsg = document.createElement('p');
                    noDataMsg.textContent = 'Sheet này không có dữ liệu.';
                    sheetContent.appendChild(noDataMsg);
                } else {
                    const table = document.createElement('table');
                    table.classList.add('excel-table');
                    const tbody = document.createElement('tbody');

                    // Header
                    const headerRow = document.createElement('tr');
                    jsonData[0].forEach(function (cellData, index) {
                        const th = document.createElement('th');
                        th.textContent = cellData !== undefined ? cellData : '';
                        th.style.width = '150px';  // Căn chỉnh cột theo chiều rộng hợp lý
                        headerRow.appendChild(th);
                    });
                    table.appendChild(headerRow);

                    // Dữ liệu bảng
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
            });

            // Cập nhật dữ liệu vào localStorage với key là fileID
            localStorage.setItem(fileID, JSON.stringify(existingData));

            // Lưu vào biến global để sử dụng sau này
            workbooks[fileID] = workbook;
            fileIDs[fileName] = fileID;  // Lưu tên file với fileID
        };

        reader.onerror = function (ex) {
            console.error('Error reading file', ex);
            alert('Đã xảy ra lỗi khi đọc tệp.');
        };

        reader.readAsBinaryString(file);
    });
}


function saveDataToLocalStorage(fileID, sheetName, updatedRow) {
    // Lấy dữ liệu hiện có từ localStorage
    const localData = JSON.parse(localStorage.getItem(fileID));

    if (!localData) {
        console.error(`Workbook with fileID ${fileID} does not exist in localStorage.`);
        return;
    }

    // Kiểm tra xem sheet đã tồn tại hay chưa
    if (!localData[sheetName]) {
        console.error(`Sheet with name ${sheetName} does not exist in file ${fileID}.`);
        return;
    }

    // Thêm dòng dữ liệu mới vào sheet
    localData[sheetName].push(updatedRow);

    // Cập nhật lại dữ liệu trong localStorage
    localStorage.setItem(fileID, JSON.stringify(localData));

    console.log(`Data for sheet ${sheetName} in file ${fileID} has been updatedqeq.`);
}


// Hàm hiển thị popup để thêm dữ liệu (cập nhật để gọi saveDataToLocalStorage)
function showAddDataPopup(fileID, sheetName, dataRow) {
    const popup = document.getElementById('popup1');
    const overlay = document.getElementById('overlay');
    const form = document.getElementById('data-form');
    const saveDataButton = document.getElementById('save-data');
    const closePopupButton = document.getElementById('close-popup');

    // Hiển thị popup và lớp nền mờ
    popup.style.display = 'block';
    overlay.style.display = 'block';

    setTimeout(() => {
        if (!form) {
            console.error('Form element not found in the popup');
            return;
        }

        form.innerHTML = '';  // Xóa nội dung cũ
        const workbook = workbooks[fileID];  // Lấy đúng workbook từ file ID
        const worksheet = workbook.Sheets[sheetName];  // Lấy đúng sheet từ workbook
        let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        const columnNames = jsonData[0];  // Tên các cột từ dòng đầu tiên

        // Tạo các ô nhập liệu từ dữ liệu trong sheet
        columnNames.forEach((columnName, index) => {
            const inputGroup = document.createElement('div');
            inputGroup.classList.add('input-group');

            const label = document.createElement('label');
            label.textContent = columnName; // Hiển thị tên cột từ file Excel
            inputGroup.appendChild(label);

            const input = document.createElement('input');
            input.type = 'text';
            input.name = `column_${index + 1}`;
            input.value = '';  // Giá trị mặc định là giá trị hiện tại trong dòng
            inputGroup.appendChild(input);

            form.appendChild(inputGroup);
        });

        // Khi nhấn nút "Lưu Dữ Liệu"
        saveDataButton.onclick = function () {
            const formData = new FormData(form);
            const updatedRow = [];

            formData.forEach((value) => {
                updatedRow.push(value);  // Thu thập dữ liệu từ form
            });

            // Tạo object theo cấu trúc ban đầu (dựa trên tên cột của sheet)
            const workbook = workbooks[fileID];  // Lấy đúng workbook từ file ID
            const worksheet = workbook.Sheets[sheetName];
            let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });  // Lấy dữ liệu hiện tại dưới dạng JSON

            const columnNames = jsonData[0];  // Tên các cột từ dòng đầu tiên

            // Chuyển array thành object với các key là tên cột
            const newDataObject = {};
            columnNames.forEach((colName, index) => {
                newDataObject[colName] = updatedRow[index] !== undefined ? updatedRow[index] : null;  // Gán dữ liệu tương ứng với cột
            });

            console.log('Dữ liệu nhập vào (dưới dạng object):', newDataObject);

            // Cập nhật dữ liệu sheet với dòng mới dưới dạng object
            jsonData.push(newDataObject);

            // Cần chuyển jsonData (mảng các object) thành array of arrays để sử dụng với aoa_to_sheet
            const updatedDataArray = [columnNames];  // Header
            jsonData.slice(1).forEach(rowObject => {  // Bỏ qua header khi chuyển đổi
                const rowArray = columnNames.map(col => rowObject[col]);  // Tạo mảng từ object
                updatedDataArray.push(rowArray);  // Thêm vào mảng các dòng
            });

            // Cập nhật lại sheet với dữ liệu mới dưới dạng array of arrays
            const updatedSheet = XLSX.utils.aoa_to_sheet(updatedDataArray);
            workbook.Sheets[sheetName] = updatedSheet;

            // Cập nhật bảng HTML (thêm dòng mới)
            const table = document.querySelector(`.sheet-container[data-file-id="${fileID}"] .excel-table tbody`);
            const newRow = document.createElement('tr');
            updatedRow.forEach(function (cellData) {
                const td = document.createElement('td');
                td.textContent = cellData !== undefined ? cellData : '';
                newRow.appendChild(td);
            });
            table.appendChild(newRow);  // Thêm dòng mới vào bảng HTML

            // Lưu dữ liệu vào localStorage theo đúng cấu trúc
            saveDataToLocalStorage(fileID, sheetName, newDataObject);

            // Ẩn popup và lớp nền
            popup.style.display = 'none';
            overlay.style.display = 'none';
        };



        // Khi đóng popup
        closePopupButton.onclick = function () {
            popup.style.display = 'none';
            overlay.style.display = 'none';
        };
    }, 0);
}

// Hàm xóa dữ liệu từ localStorage khi tải lại trang
window.addEventListener('DOMContentLoaded', function () {
    // Lặp qua các mục trong localStorage và xóa tất cả các mục có key bắt đầu bằng "file_"
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.startsWith('file_')) {
            localStorage.removeItem(key);
            i--; // Giảm i vì localStorage đã thay đổi kích thước sau khi xóa phần tử
        }
    }

    console.log('Tất cả dữ liệu liên quan đến các file đã bị xóa khỏi localStorage.');
});


function calculateColumnWidths(jsonData) {
    const colWidths = [];
    jsonData.forEach(row => {
        row.forEach((cell, index) => {
            const cellLength = cell ? cell.toString().length : 0;
            if (!colWidths[index] || cellLength > colWidths[index]) {
                colWidths[index] = cellLength;
            }
        });
    });
    return colWidths.map(width => width * 10);  // Đoạn mã này nhân với 10 để có độ rộng phù hợp
}