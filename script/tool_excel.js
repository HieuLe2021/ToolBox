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
            const workbook = XLSX.read(data, { type: 'binary' });  // Đọc file Excel và lưu vào một workbook riêng

            workbooks[fileID] = workbook;  // Lưu workbook theo ID file
            fileIDs[fileName] = fileID;  // Lưu tên file tương ứng với ID

            const sheetNames = workbook.SheetNames;

            sheetNames.forEach(function (sheetName) {
                if (sheetName === 'hiddenSheet') {
                    return;  // Bỏ qua sheet có tên là "hiddenSheet"
                }

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

                const worksheet = workbook.Sheets[sheetName];
                let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                // Loại bỏ các cột không cần thiết (ID, Row Checksum, Modified On)
                const columnsToRemove = ['(Do Not Modify) ID', '(Do Not Modify) Row Checksum', '(Do Not Modify) Modified On'];
                jsonData = jsonData.map(row => {
                    return row.filter((cell, index) => {
                        return !columnsToRemove.includes(jsonData[0][index]);  // Kiểm tra tên cột
                    });
                });

                const columns = jsonData[0];  // Lấy tên các cột

                // Format lại dữ liệu JSON thành array of objects
                const formattedRows = jsonData.slice(1).map(row => {
                    const rowObject = {};
                    row.forEach((cellData, index) => {
                        rowObject[columns[index]] = cellData;
                    });
                    return rowObject;
                });

                // Lưu vào localStorage
                const storedData = JSON.parse(localStorage.getItem(fileID)) || [];
                if (!storedData.length) {
                    localStorage.setItem(fileID, JSON.stringify(formattedRows));
                }

                // Tính toán chiều rộng cột dựa trên dữ liệu
                const colWidths = calculateColumnWidths(jsonData);

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

                const sheetContent = document.createElement('div');
                sheetContent.classList.add('sheet-content');
                sheetContent.style.display = 'none';

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
                        th.style.width = `${colWidths[index]}px`;  // Căn chỉnh cột theo chiều rộng tính toán
                        headerRow.appendChild(th);
                    });
                    table.appendChild(headerRow);

                    // Dữ liệu bảng
                    for (let i = 1; i < jsonData.length; i++) {
                        const rowData = jsonData[i];
                        const row = document.createElement('tr');
                        rowData.forEach(function (cellData, index) {
                            const td = document.createElement('td');
                            td.textContent = cellData !== undefined ? cellData : '';
                            td.style.width = `${colWidths[index]}px`;  // Căn chỉnh cột theo chiều rộng tính toán
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
        };

        reader.onerror = function (ex) {
            console.error('Error reading file', ex);
            alert('Đã xảy ra lỗi khi đọc tệp.');
        };

        reader.readAsBinaryString(file);
    });
}


// Hàm cập nhật workbook và lưu vào localStorage khi người dùng nhập dữ liệu
// Hàm lưu dữ liệu vào localStorage (được cập nhật để hỗ trợ thêm hoặc cập nhật)
function saveDataToLocalStorage(workbooks, fileID, sheetName, updatedRow) {
    // Lấy dữ liệu từ localStorage (nếu có)
    let storedData = localStorage.getItem('excelData');
    let dataToSave = storedData ? JSON.parse(storedData) : { data: {}, timestamp: '' };

    // Lấy workbook và sheet đúng từ workbooks
    const workbook = workbooks[fileID];
    const worksheet = workbook.Sheets[sheetName];
    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const columns = jsonData[0]; // Tên các cột
    const rows = jsonData.slice(1); // Các dòng dữ liệu đã có

    // Tạo đối tượng chứa các dữ liệu dưới dạng key-value (mỗi dòng dữ liệu là một đối tượng)
    const formattedRows = rows.map(row => {
        const rowObject = {};
        row.forEach((cellData, index) => {
            rowObject[columns[index]] = cellData;
        });
        return rowObject;
    });

    // Nếu sheet đã có dữ liệu trong localStorage thì cập nhật, nếu chưa có thì thêm mới
    if (dataToSave.data[sheetName]) {
        console.log(`Sheet ${sheetName} đã tồn tại trong localStorage, đang cập nhật...`);
        dataToSave.data[sheetName].push(...formattedRows); // Cập nhật dữ liệu bằng cách thêm dòng mới
    } else {
        console.log(`Sheet ${sheetName} chưa có trong localStorage, thêm mới...`);
        dataToSave.data[sheetName] = formattedRows; // Thêm mới dữ liệu
    }

    // Lưu lại thời gian cập nhật
    dataToSave.timestamp = new Date().toISOString();

    // Lưu dữ liệu đã cập nhật vào localStorage
    localStorage.setItem('excelData', JSON.stringify(dataToSave));

    console.log('Dữ liệu đã được cập nhật vào localStorage với timestamp:', dataToSave.timestamp);
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
            formData.forEach((value, key) => {
                updatedRow.push(value);
            });
            console.log('Dữ liệu nhập vào:', updatedRow);

            // Cập nhật workbook và sheet đúng
            const workbook = workbooks[fileID];  // Lấy đúng workbook từ file ID
            const worksheet = workbook.Sheets[sheetName];  // Lấy đúng sheet từ workbook

            let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Loại bỏ các cột không cần thiết (ID, Row Checksum, Modified On)
            const columnsToRemove = ['(Do Not Modify) ID', '(Do Not Modify) Row Checksum', '(Do Not Modify) Modified On'];
            jsonData = jsonData.map(row => {
                return row.filter((cell, index) => {
                    return !columnsToRemove.includes(jsonData[0][index]);  // Kiểm tra tên cột
                });
            });

            jsonData.push(updatedRow);  // Thêm dòng mới vào dữ liệu hiện tại

            // Cập nhật sheet với dữ liệu mới
            const updatedSheet = XLSX.utils.aoa_to_sheet(jsonData);
            workbook.Sheets[sheetName] = updatedSheet;

            // Ẩn popup và lớp nền
            popup.style.display = 'none';
            overlay.style.display = 'none';

            // Cập nhật bảng HTML
            const table = document.querySelector(`.sheet-container[data-file-id="${fileID}"] .excel-table tbody`);
            const newRow = document.createElement('tr');
            updatedRow.forEach(function (cellData) {
                const td = document.createElement('td');
                td.textContent = cellData !== undefined ? cellData : '';
                newRow.appendChild(td);
            });
            table.appendChild(newRow);  // Thêm dòng mới vào bảng HTML

            // Lưu dữ liệu cập nhật vào localStorage
            saveDataToLocalStorage(workbooks, fileID, sheetName, updatedRow);
        };

        // Khi đóng popup
        closePopupButton.onclick = function () {
            popup.style.display = 'none';
            overlay.style.display = 'none';
        };
    }, 0);
}

// Hàm lưu hoặc cập nhật dữ liệu vào localStorage
// Function to save or update data in localStorage
function saveDataToLocalStorage(fileID, sheetName, updatedRow) {
    // Get current localStorage data
    const localData = JSON.parse(localStorage.getItem(fileID)); // Use fileID directly
    if (!localData) {
        console.error(`Workbook with fileID ${fileID} does not exist in localStorage.`);
        return;
    }

    const workbook = localData; // Assume this is already in the correct format

    // Check if the sheet exists in the workbook
    if (!workbook[sheetName]) {
        console.error(`Sheet with name ${sheetName} does not exist in workbook ${fileID}.`);
        return;
    }

    // Update the data in the existing sheet
    workbook[sheetName].push(updatedRow);

    // Save the updated workbook back to localStorage
    localStorage.setItem(fileID, JSON.stringify(workbook));

    console.log(`Data for sheet ${sheetName} in file ${fileID} has been updated.`);
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