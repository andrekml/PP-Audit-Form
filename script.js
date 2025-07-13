// Initialize entries array from localStorage or create empty array
let entries = JSON.parse(localStorage.getItem('workorderEntries')) || [];
let currentEditId = null;

// Validate required fields
function validateForm() {
    let isValid = true;
    const requiredFields = [
        'transformerNo', 
        'latitude', 
        'longitude'
    ];

    requiredFields.forEach(field => {
        const element = document.getElementById(field);
        const errorElement = document.getElementById(`${field}-error`);
        if (!element.value.trim()) {
            element.classList.add('error-border');
            errorElement.style.display = 'block';
            isValid = false;
        } else {
            element.classList.remove('error-border');
            errorElement.style.display = 'none';
        }
    });

    return isValid;
}

// Save entry to local storage
document.getElementById('saveEntry').addEventListener('click', function() {
    if (!validateForm()) {
        alert('Please fill in all required fields (marked with *)');
        return;
    }

    const formData = collectFormData();
    
    if (currentEditId !== null) {
        // Update existing entry
        const index = entries.findIndex(entry => entry.id === currentEditId);
        if (index !== -1) {
            entries[index] = formData;
            alert('Entry updated successfully!');
        }
    } else {
        // Add new entry
        formData.id = Date.now().toString(); // Simple unique ID
        entries.push(formData);
        alert('Entry saved successfully!');
    }
    
    // Save to localStorage
    localStorage.setItem('workorderEntries', JSON.stringify(entries));
    
    // Reset form
    document.getElementById('dataForm').reset();
    document.getElementById('entryId').value = '';
    currentEditId = null;
});

// View saved entries
document.getElementById('viewEntries').addEventListener('click', function() {
    displayEntries();
    document.getElementById('entriesModal').style.display = 'block';
});

// Close modal
document.querySelector('.close').addEventListener('click', function() {
    document.getElementById('entriesModal').style.display = 'none';
});

// Search functionality
document.getElementById('searchInput').addEventListener('input', function(e) {
    const searchTerm = e.target.value.toLowerCase();
    const filteredEntries = entries.filter(entry => 
        (entry.workorderNo && entry.workorderNo.toLowerCase().includes(searchTerm)) ||
        (entry.date && entry.date.toLowerCase().includes(searchTerm)) ||
        (entry.contractorName && entry.contractorName.toLowerCase().includes(searchTerm)) ||
        (entry.transformerNo && entry.transformerNo.toLowerCase().includes(searchTerm))
    );
    displayEntries(filteredEntries);
});

// Export to Excel with auto-capitalization and professional formatting
document.getElementById('saveAsExcel').addEventListener('click', function() {
    if (entries.length === 0) {
        alert('No entries to export!');
        return;
    }

    // Load SheetJS library if not already loaded
    if (typeof XLSX === 'undefined') {
        alert('Excel export feature is still loading. Please try again in a moment.');
        return;
    }

    // Create Excel workbook with multiple sheets
    const workbook = XLSX.utils.book_new();
    
    // Process data with auto-capitalization
    const mainData = entries.map(entry => {
        const formattedEntry = {};
        Object.keys(entry).forEach(key => {
            if (key !== 'id') {
                let value = entry[key] || '';
                
                // Auto-capitalize first letter of each word for text fields
                if (typeof value === 'string' && value.length > 0) {
                    // Skip capitalization for these fields:
                    const skipCapitalize = ['workorderNo', 'meterNo', 'newMeterNo', 'tariffCode', 
                                          'supplyGroupCode', 'tamperNotNo', 'mmfNo', 'sealNo',
                                          'phoneCode', 'phoneNo', 'idNo', 'date', 'latitude',
                                          'longitude', 'transformerNo'];
                                          
                    if (!skipCapitalize.includes(key)) {
                        value = value.toLowerCase()
                            .split(' ')
                            .map(word => word.charAt(0).toUpperCase() + word.slice(1))
                            .join(' ');
                    }
                }
                
                // Format dates properly
                if (key === 'date' && value) {
                    formattedEntry[key] = new Date(value);
                } else {
                    formattedEntry[key] = value;
                }
            }
        });
        return formattedEntry;
    });

    // Create worksheet with auto-width columns
    const mainWorksheet = XLSX.utils.json_to_sheet(mainData);
    
    // Capitalize headers in the worksheet
    const range = XLSX.utils.decode_range(mainWorksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_col(C) + "1"; // First row is headers
        if (!mainWorksheet[address]) continue;
        
        // Convert camelCase to proper capitalization with spaces
        let header = mainWorksheet[address].v;
        header = header
            .replace(/([A-Z])/g, ' $1') // Add space before capital letters
            .replace(/^./, str => str.toUpperCase()) // Capitalize first letter
            .trim();
        
        mainWorksheet[address].v = header;
        mainWorksheet[address].t = 's'; // Ensure it's stored as string
    }
    
    // Set column widths based on content
    const colWidths = [];
    const headerKeys = Object.keys(mainData[0]);
    headerKeys.forEach(key => {
        const maxLength = Math.max(
            key.replace(/([A-Z])/g, ' $1').length + 1, // Account for added spaces
            ...mainData.map(row => (row[key] ? row[key].toString().length : 0))
        );
        colWidths.push({ wch: Math.min(Math.max(maxLength, 10), 30) });
    });
    
    mainWorksheet['!cols'] = colWidths;
    
    // Add worksheet to workbook with capitalized sheet name
    XLSX.utils.book_append_sheet(workbook, mainWorksheet, "Workorders");
    
    // Summary sheet with improved formatting
    const summaryData = [
        ["WORKORDER DATA REPORT", "", "", "", ""],
        ["", "", "", "", ""],
        ["Report Summary", "", "", "", ""],
        ["Total Entries", entries.length, "", "", ""],
        ["Generated On", new Date().toLocaleString(), "", "", ""],
        ["", "", "", "", ""],
        ["Field Name", "Non-empty Count", "Sample Value", "", ""],
        ...Object.keys(entries[0])
            .filter(key => key !== 'id')
            .map(key => {
                const nonEmptyEntries = entries.filter(entry => entry[key] && entry[key].toString().trim() !== '');
                const sampleValue = nonEmptyEntries.length > 0 ? nonEmptyEntries[0][key] : '';
                return [
                    key.replace(/([A-Z])/g, ' $1') // Add space before capital letters
                       .replace(/^./, str => str.toUpperCase()) // Capitalize first letter
                       .trim(),
                    nonEmptyEntries.length,
                    sampleValue,
                    "",
                    ""
                ];
            })
    ];
    
    const summaryWorksheet = XLSX.utils.aoa_to_sheet(summaryData);
    
    // Format summary sheet
    summaryWorksheet['!cols'] = [
        { wch: 25 }, { wch: 15 }, { wch: 25 }, { wch: 10 }, { wch: 10 }
    ];
    
    // Add header styling
    const headerStyle = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "3498db" } } };
    for (let i = 0; i < 7; i++) {
        if (!summaryWorksheet[`A${i+1}`]) continue;
        if (!summaryWorksheet[`A${i+1}`].s) summaryWorksheet[`A${i+1}`].s = {};
        Object.assign(summaryWorksheet[`A${i+1}`].s, headerStyle);
    }
    
    XLSX.utils.book_append_sheet(workbook, summaryWorksheet, "Summary");
    
    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array',
        cellStyles: true 
    });
    
    saveAsExcelFile(excelBuffer, `Workorder_Data_${formatDate(new Date())}.xlsx`);
});

// Helper function to save Excel file with auto-capitalization
function saveAsExcelFile(buffer, fileName) {
    // Auto-capitalize filename
    fileName = fileName
        .split('_')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
        .join(' ')
        .replace('.xlsx', '') + '.xlsx';
    
    const data = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    const url = URL.createObjectURL(data);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Format date for filename
function formatDate(date) {
    const d = new Date(date);
    return [
        d.getFullYear(),
        (d.getMonth() + 1).toString().padStart(2, '0'),
        d.getDate().toString().padStart(2, '0')
    ].join('-');
}

// Reset form
document.getElementById('resetForm').addEventListener('click', function() {
    if(confirm('Are you sure you want to reset all fields?')) {
        document.getElementById('dataForm').reset();
        document.getElementById('entryId').value = '';
        currentEditId = null;
        
        // Clear error states
        document.querySelectorAll('.error-border').forEach(el => {
            el.classList.remove('error-border');
        });
        document.querySelectorAll('.error-message').forEach(el => {
            el.style.display = 'none';
        });
    }
});

// Helper function to collect form data
function collectFormData() {
    const formData = {
        id: document.getElementById('entryId').value || Date.now().toString(),
        workorderNo: document.getElementById('workorderNo').value,
        zone: document.getElementById('zone').value,
        sector: document.getElementById('sector').value,
        cnc: document.getElementById('cnc').value,
        contractorName: document.getElementById('contractorName').value,
        date: document.getElementById('date').value,
        feederName: document.getElementById('feederName').value,
        townshipName: document.getElementById('townshipName').value,
        transformerNo: document.getElementById('transformerNo').value,
        standNo: document.getElementById('standNo').value,
        installationNo: document.getElementById('installationNo').value,
        latitude: document.getElementById('latitude').value,
        longitude: document.getElementById('longitude').value,
        access: document.getElementById('access').value,
        atHome: document.getElementById('atHome').value,
        connected: document.getElementById('connected').value,
        abandoned: document.getElementById('abandoned').value,
        vandalised: document.getElementById('vandalised').value,
        illegal: document.getElementById('illegal').value,
        owner: document.getElementById('owner').value,
        surname: document.getElementById('surname').value,
        name: document.getElementById('name').value,
        idNo: document.getElementById('idNo').value,
        phoneCode: document.getElementById('phoneCode').value,
        phoneNo: document.getElementById('phoneNo').value,
        normalVendor: document.getElementById('normalVendor').value,
        mostRecentVendor: document.getElementById('mostRecentVendor').value,
        meterNo: document.getElementById('meterNo').value,
        split: document.getElementById('split').value,
        commonBaseMeter: document.getElementById('commonBaseMeter').value,
        edEcu: document.getElementById('edEcu').value,
        meterType: document.getElementById('meterType').value,
        tariffCode: document.getElementById('tariffCode').value,
        ampLimit: document.getElementById('ampLimit').value,
        supplyGroupCode: document.getElementById('supplyGroupCode').value,
        magCardKeypad: document.getElementById('magCardKeypad').value,
        stsProp: document.getElementById('stsProp').value,
        tampered: document.getElementById('tampered').value,
        tamperMethod: document.getElementById('tamperMethod').value,
        removed: document.getElementById('removed').value,
        tamperNotNo: document.getElementById('tamperNotNo').value,
        cardTrip: document.getElementById('cardTrip').value,
        elTest: document.getElementById('elTest').value,
        meterFaulty: document.getElementById('meterFaulty').value,
        replaced: document.getElementById('replaced').value,
        mmfNo: document.getElementById('mmfNo').value,
        newMeterNo: document.getElementById('newMeterNo').value,
        newMeterType: document.getElementById('newMeterType').value,
        newMeterAmpLimit: document.getElementById('newMeterAmpLimit').value,
        newMeterCardTrip: document.getElementById('newMeterCardTrip').value,
        newMeterElTest: document.getElementById('newMeterElTest').value,
        sealed: document.getElementById('sealed').value,
        sealNo: document.getElementById('sealNo').value,
        remarks: document.getElementById('remarks').value
    };
    
    return formData;
}

// Helper function to display entries in the modal
function displayEntries(entriesToDisplay = null) {
    const entriesToShow = entriesToDisplay || entries;
    const tableBody = document.getElementById('entriesTableBody');
    tableBody.innerHTML = '';

    if (entriesToShow.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" style="text-align: center;">No entries found</td>`;
        tableBody.appendChild(row);
        return;
    }

    entriesToShow.forEach(entry => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${entry.workorderNo || 'N/A'}</td>
            <td>${entry.date || ''}</td>
            <td>${entry.contractorName || ''}</td>
            <td>${entry.transformerNo || ''}</td>
            <td>${entry.latitude ? entry.latitude + ', ' + entry.longitude : ''}</td>
            <td>
                <button class="action-btn edit-btn" data-id="${entry.id}">Edit</button>
                <button class="action-btn delete-btn" data-id="${entry.id}">Delete</button>
                <button class="action-btn export-btn" data-id="${entry.id}">Export</button>
            </td>
        `;
        tableBody.appendChild(row);
    });

    // Add event listeners to action buttons
    document.querySelectorAll('.edit-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            editEntry(this.getAttribute('data-id'));
        });
    });

    document.querySelectorAll('.delete-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            deleteEntry(this.getAttribute('data-id'));
        });
    });

    document.querySelectorAll('.export-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            exportSingleEntry(this.getAttribute('data-id'));
        });
    });
}

// Helper function to edit an entry
function editEntry(id) {
    const entry = entries.find(e => e.id === id);
    if (!entry) return;

    currentEditId = id;
    document.getElementById('entryId').value = id;

    // Populate form fields
    Object.keys(entry).forEach(key => {
        if (key !== 'id' && document.getElementById(key)) {
            document.getElementById(key).value = entry[key] || '';
        }
    });

    // Close modal
    document.getElementById('entriesModal').style.display = 'none';
    
    // Scroll to top
    window.scrollTo(0, 0);
}

// Helper function to delete an entry
function deleteEntry(id) {
    if (!confirm('Are you sure you want to delete this entry?')) return;

    entries = entries.filter(entry => entry.id !== id);
    localStorage.setItem('workorderEntries', JSON.stringify(entries));
    displayEntries();
}

// Helper function to export a single entry
function exportSingleEntry(id) {
    const entry = entries.find(e => e.id === id);
    if (!entry) return;

    // Create a workbook with just this entry
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet([entry]);
    
    // Capitalize headers in the worksheet
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_col(C) + "1"; // First row is headers
        if (!worksheet[address]) continue;
        
        // Convert camelCase to proper capitalization with spaces
        let header = worksheet[address].v;
        header = header
            .replace(/([A-Z])/g, ' $1') // Add space before capital letters
            .replace(/^./, str => str.toUpperCase()) // Capitalize first letter
            .trim();
        
        worksheet[address].v = header;
        worksheet[address].t = 's'; // Ensure it's stored as string
    }
    
    XLSX.utils.book_append_sheet(workbook, worksheet, "Workorder");

    // Generate Excel file
    const excelBuffer = XLSX.write(workbook, { 
        bookType: 'xlsx', 
        type: 'array',
        cellStyles: true 
    });
    
    saveAsExcelFile(excelBuffer, `Workorder_${entry.workorderNo || id}_${formatDate(new Date())}.xlsx`);
}

// Load SheetJS library for Excel export
const script = document.createElement('script');
script.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
script.onload = function() {
    console.log('SheetJS library loaded');
};
document.head.appendChild(script);