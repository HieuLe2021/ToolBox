// main.js

// Hàm để load sidebar
function loadSidebar(activeLinkId) {
    fetch('../component/sidebar.html')
        .then(response => response.text())
        .then(data => {
            // Chèn sidebar vào placeholder
            document.getElementById('sidebar-placeholder').outerHTML = data;

            // Thêm lớp 'active' cho link hiện tại
            if (activeLinkId) {
                const activeLink = document.getElementById(activeLinkId);
                if (activeLink) {
                    activeLink.classList.add('active');
                }
            }
        })
        .catch(error => {
            console.error('Error loading sidebar:', error);
        });
}