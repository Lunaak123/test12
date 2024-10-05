function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

const sheetName = getQueryParam('sheetName');
const fileUrl = getQueryParam('fileUrl');

(async () => {
    if (!fileUrl || !sheetName) {
        alert("Invalid sheet data.");
        return;
    }

    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            alert("Sheet not found.");
            return;
        }

        // Convert sheet to HTML table and display
        const html = XLSX.utils.sheet_to_html(sheet);
        const sheetContentDiv = document.getElementById('sheet-content');
        sheetContentDiv.innerHTML = html;

        document.getElementById('apply-operation').addEventListener('click', () => {
            const primaryCol = document.getElementById('primary-column').value.toUpperCase();
            const operationCols = document.getElementById('operation-columns').value.toUpperCase().split(',');
            const operationType = document.getElementById('operation-type').value;

            if (!primaryCol || !operationCols.length) {
                alert("Please enter valid column names.");
                return;
            }

            const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            const header = data[0];
            const primaryIndex = header.indexOf(primaryCol);

            if (primaryIndex === -1) {
                alert("Primary column not found.");
                return;
            }

            const filteredData = data.filter((row, index) => {
                if (index === 0) return true; // Keep header

                const primaryCell = row[primaryIndex];

                // Check for the specified operation columns
                let allOperationsNull = true;
                let anyOperationNotNull = false;

                operationCols.forEach(col => {
                    const operationIndex = header.indexOf(col);
                    if (operationIndex !== -1) {
                        const operationCell = row[operationIndex];
                        if (operationCell) anyOperationNotNull = true;
                        if (!operationCell) allOperationsNull = allOperationsNull && true;
                    }
                });

                if (operationType === 'and') {
                    return primaryCell !== null && primaryCell !== "" && !allOperationsNull;
                } else {
                    return primaryCell !== null && primaryCell !== "" && (anyOperationNotNull);
                }
            });

            // Update the sheet content
            const newHtml = XLSX.utils.sheet_to_html(XLSX.utils.aoa_to_sheet(filteredData));
            sheetContentDiv.innerHTML = newHtml;
        });
    } catch (error) {
        console.error("Error loading Excel file:", error);
        alert("Failed to load the Excel sheet. Please check the URL and try again.");
    }
})();
