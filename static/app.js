// Global State Variables
let currentPrinters = [];
let defaultPrinter = "";
let printAgentConnected = false;
let currentGeneratedQRs = [];
let previewCurrentPage = 1;
let previewTotalPages = 1;

let selectedPdfId = null;
let selectedPdfPrinted = false;
let selectedPdfFileName = "";

// ---------------------------------------------------------
// Initialization on Document Load
// ---------------------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
    initTabs();
    checkPrintAgentStatus();
    loadModels();
    loadPdfGrid();
    initReports();
    initUserManagement();
    
    // Auto-reconnect to print agent every 5 seconds if disconnected
    setInterval(() => {
        if (!printAgentConnected) {
            checkPrintAgentStatus();
        }
    }, 5000);
});

// Display global alert messages
function showGlobalAlert(message, type = "info", duration = 4000) {
    const alert = document.getElementById("global-alert");
    alert.className = `alert alert-${type}`;
    alert.innerText = message;
    alert.classList.remove("hide");
    
    setTimeout(() => {
        alert.classList.add("hide");
    }, duration);
}

// ---------------------------------------------------------
// Print Agent Connectivity Checks
// ---------------------------------------------------------
async function checkPrintAgentStatus() {
    const dot = document.getElementById("agent-status-dot");
    const text = document.getElementById("agent-status-text");
    const dropdowns = document.querySelectorAll(".printer-dropdown");
    
    try {
        const response = await fetch("http://localhost:5001/printers", { method: "GET" });
        if (!response.ok) throw new Error("Agent error");
        
        const data = await response.json();
        currentPrinters = data.printers;
        defaultPrinter = data.default;
        printAgentConnected = true;
        
        dot.className = "status-dot green";
        text.innerText = "Print Agent: Online";
        
        // Populate dropdown selections
        dropdowns.forEach(dd => {
            dd.innerHTML = "";
            currentPrinters.forEach(prn => {
                const opt = document.createElement("option");
                opt.value = prn;
                opt.text = prn;
                if (prn === defaultPrinter) {
                    opt.selected = true;
                }
                dd.appendChild(opt);
            });
        });
        
        // Enable sequential print if we have generated QR codes
        if (currentGeneratedQRs.length > 0) {
            document.getElementById("btn-print-sequential").disabled = false;
        }
        // Enable print PDF button if a PDF is selected
        if (selectedPdfId) {
            document.getElementById("btn-print-pdf").disabled = false;
        }
        
    } catch (err) {
        printAgentConnected = false;
        dot.className = "status-dot red";
        text.innerText = "Print Agent: Offline";
        
        dropdowns.forEach(dd => {
            dd.innerHTML = '<option value="">-- No Print Agent Connected --</option>';
        });
        
        document.getElementById("btn-print-sequential").disabled = true;
        document.getElementById("btn-print-pdf").disabled = true;
    }
}

// Load models into dropdowns
async function loadModels() {
    try {
        const response = await fetch("/api/items");
        const items = await response.json();
        
        const seqSelect = document.getElementById("seq-item-select");
        const pdfSelect = document.getElementById("pdf-item-select");
        
        // Populate normal PDF import dropdown
        const optHtml = ['<option value="">-- Select Model --</option>']
            .concat(items.map(item => `<option value="${item.CodeStr}">${item.CodeStr} - ${item.Name}</option>`))
            .join("");
        pdfSelect.innerHTML = optHtml;
        
        // Set up the custom searchable dropdown for the sequential tab
        setupSearchableDropdown(items);
        
        // Load last serial when sequential item select changes
        seqSelect.addEventListener("change", updateLastSerialDisplay);
        
    } catch (err) {
        showGlobalAlert("Failed to load models from database server.", "danger");
    }
}

// Custom Searchable Dropdown Layout logic (Android-contacts-like real-time filtering)
function setupSearchableDropdown(items) {
    const searchInput = document.getElementById("seq-item-search");
    const hiddenSelect = document.getElementById("seq-item-select");
    const optionsList = document.getElementById("seq-item-options-list");
    let highlightedIndex = -1;
    
    function renderOptions(filterText = "") {
        optionsList.innerHTML = "";
        highlightedIndex = -1;
        
        const filtered = items.filter(item => {
            const searchStr = `${item.CodeStr} ${item.Name}`.toLowerCase();
            return searchStr.includes(filterText.toLowerCase());
        });
        
        if (filtered.length === 0) {
            const noResults = document.createElement("div");
            noResults.className = "searchable-select-no-results";
            noResults.textContent = "No matching models found";
            optionsList.appendChild(noResults);
            return;
        }
        
        filtered.forEach((item, index) => {
            const div = document.createElement("div");
            div.className = "searchable-select-item";
            div.dataset.value = item.CodeStr;
            div.dataset.index = index;
            div.textContent = `${item.CodeStr} - ${item.Name}`;
            
            // Use mousedown to prevent blur race condition
            div.addEventListener("mousedown", (e) => {
                e.preventDefault();
                selectItem(item);
            });
            
            optionsList.appendChild(div);
        });
    }
    
    function selectItem(item) {
        searchInput.value = `${item.CodeStr} - ${item.Name}`;
        hiddenSelect.value = item.CodeStr;
        optionsList.classList.add("hide");
        
        // Trigger change event to load last serial automatically
        const event = new Event('change');
        hiddenSelect.dispatchEvent(event);
    }
    
    // Toggle options list when input is focused
    searchInput.addEventListener("focus", () => {
        renderOptions(searchInput.value);
        optionsList.classList.remove("hide");
    });
    
    // Close options list when input loses focus
    searchInput.addEventListener("blur", () => {
        setTimeout(() => {
            optionsList.classList.add("hide");
        }, 150);
    });
    
    // Filter list on keyup
    searchInput.addEventListener("input", () => {
        if (!searchInput.value.trim()) {
            hiddenSelect.value = "";
            const event = new Event('change');
            hiddenSelect.dispatchEvent(event);
        }
        renderOptions(searchInput.value);
        optionsList.classList.remove("hide");
    });
    
    // Keyboard navigation (Arrow keys + Enter + Escape)
    searchInput.addEventListener("keydown", (e) => {
        const itemElems = optionsList.querySelectorAll(".searchable-select-item");
        if (itemElems.length === 0) return;
        
        if (e.key === "ArrowDown") {
            e.preventDefault();
            highlightedIndex = (highlightedIndex + 1) % itemElems.length;
            updateHighlight(itemElems);
        } else if (e.key === "ArrowUp") {
            e.preventDefault();
            highlightedIndex = (highlightedIndex - 1 + itemElems.length) % itemElems.length;
            updateHighlight(itemElems);
        } else if (e.key === "Enter") {
            e.preventDefault();
            if (highlightedIndex >= 0 && highlightedIndex < itemElems.length) {
                const selectedValue = itemElems[highlightedIndex].dataset.value;
                const foundItem = items.find(item => item.CodeStr === selectedValue);
                if (foundItem) {
                    selectItem(foundItem);
                    searchInput.blur();
                }
            }
        } else if (e.key === "Escape") {
            optionsList.classList.add("hide");
            searchInput.blur();
        }
    });
    
    function updateHighlight(itemElems) {
        itemElems.forEach((item, idx) => {
            if (idx === highlightedIndex) {
                item.classList.add("highlighted");
                item.scrollIntoView({ block: "nearest" });
            } else {
                item.classList.remove("highlighted");
            }
        });
    }
}

// ---------------------------------------------------------
// Tab Panel Switcher
// ---------------------------------------------------------
function initTabs() {
    const menuItems = document.querySelectorAll(".menu-item");
    const panes = document.querySelectorAll(".tab-pane");
    
    menuItems.forEach(item => {
        item.addEventListener("click", () => {
            // Remove active status
            menuItems.forEach(mi => mi.classList.remove("active"));
            panes.forEach(p => p.classList.remove("active"));
            
            // Set active status
            item.classList.add("active");
            const target = item.getAttribute("data-target");
            document.getElementById(target).classList.add("active");
        });
    });
}

// ---------------------------------------------------------
// Tab: Sequential Generation Logic
// ---------------------------------------------------------
async function updateLastSerialDisplay() {
    const itemCode = document.getElementById("seq-item-select").value;
    const label = document.getElementById("last-generated-serial");
    
    if (!itemCode) {
        label.innerText = "Last Generated QR: N/A";
        return;
    }
    
    label.innerText = "Last Generated QR: Loading...";
    try {
        const response = await fetch(`/api/last_serial?item_code=${encodeURIComponent(itemCode)}`);
        const data = await response.json();
        label.innerText = `Last Generated QR: ${data.last_serial}`;
    } catch (err) {
        label.innerText = "Last Generated QR: Error";
    }
}

document.getElementById("btn-generate-qrs").addEventListener("click", async () => {
    const itemCode = document.getElementById("seq-item-select").value;
    const qty = parseInt(document.getElementById("seq-qty").value, 10);
    const btn = document.getElementById("btn-generate-qrs");
    
    if (!itemCode || qty <= 0) {
        showGlobalAlert("Please select a valid model and enter quantity.", "danger");
        return;
    }
    
    btn.disabled = true;
    btn.innerText = "Generating...";
    
    try {
        const response = await fetch("/api/generate_qrs", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ item_code: itemCode, qty })
        });
        
        const result = await response.json();
        if (result.success) {
            currentGeneratedQRs = result.qrs;
            previewCurrentPage = 1;
            previewTotalPages = Math.ceil(currentGeneratedQRs.length / 4);
            
            showGlobalAlert(result.message, "success");
            updateLastSerialDisplay();
            drawPreviewPage();
            
            if (printAgentConnected) {
                document.getElementById("btn-print-sequential").disabled = false;
            }
        } else {
            showGlobalAlert(result.message || "Failed to generate QR codes.", "danger");
            currentGeneratedQRs = [];
            document.getElementById("btn-print-sequential").disabled = true;
        }
    } catch (err) {
        showGlobalAlert("Server connection error during generation.", "danger");
    } finally {
        btn.disabled = false;
        btn.innerText = "Generate QR codes";
    }
});

// Canvas preview rendering (Draws 4 labels horizontally using actual server-generated QRs)
function drawPreviewPage() {
    const canvas = document.getElementById("preview-canvas");
    const ctx = canvas.getContext("2d");
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    
    if (currentGeneratedQRs.length === 0) return;
    
    // Retrieve inputs (mm)
    const labelWidthMm = parseFloat(document.getElementById("seq-label-width").value) || 25;
    const labelHeightMm = parseFloat(document.getElementById("seq-label-height").value) || 25;
    const gapMm = parseFloat(document.getElementById("seq-gap").value) || 3;
    const topMarginMm = parseFloat(document.getElementById("seq-top-margin").value) || 5;
    
    // Convert mm to pixels at a default base scale of 4 pixels per mm
    const baseScale = 4.0;
    let w = labelWidthMm * baseScale;
    let h = labelHeightMm * baseScale;
    let gap = gapMm * baseScale;
    let topMargin = topMarginMm * baseScale;
    let leftMargin = 30.0;
    
    // Auto-scale to fit canvas width (600px)
    const totalRowWidth = 4 * w + 3 * gap;
    const maxAvailableWidth = canvas.width - 60; // 30px margins on both sides
    if (totalRowWidth > maxAvailableWidth) {
        const scale = maxAvailableWidth / totalRowWidth;
        w *= scale;
        h *= scale;
        gap *= scale;
        topMargin *= scale;
    }
    
    // Auto-scale canvas height dynamically if height is large
    if (topMargin + h + 25 > canvas.height) {
        canvas.height = topMargin + h + 35;
    } else {
        canvas.height = 250; // Reset to default
    }
    
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    const labelsPerPage = 4;
    
    for (let i = 0; i < labelsPerPage; i++) {
        const idx = (previewCurrentPage - 1) * labelsPerPage + i;
        if (idx >= currentGeneratedQRs.length) break;
        
        const qrValue = currentGeneratedQRs[idx];
        const parts = qrValue.split("|");
        const labelText = parts.length >= 7 ? `${parts[2].trim()}-${parts[6].trim()}` : qrValue;
        
        const x = leftMargin + i * (w + gap);
        
        // 1. Draw outer label border
        ctx.strokeStyle = "#94a3b8";
        ctx.lineWidth = 1;
        ctx.strokeRect(x, topMargin, w, h);
        
        // 2. Draw actual QR code from server
        const img = new Image();
        img.src = `/api/qr_image?text=${encodeURIComponent(qrValue)}`;
        img.onload = () => {
            if (w > h * 1.3) {
                // Rectangular Layout (Side-by-Side)
                const qrSize = h - 6;
                ctx.drawImage(img, x + 3, topMargin + 3, qrSize, qrSize);
                
                ctx.fillStyle = "#000000";
                ctx.font = `bold ${Math.max(8, Math.min(11, h * 0.35))}px Outfit, Arial, sans-serif`;
                ctx.textAlign = "left";
                ctx.textBaseline = "middle";
                ctx.fillText(labelText, x + qrSize + 8, topMargin + h / 2);
            } else {
                // Square/Vertical Layout (Stacked)
                const qrSize = Math.min(w, h) - 20;
                const qrX = x + (w - qrSize) / 2;
                const qrY = topMargin + (h - qrSize) / 2 - 5;
                ctx.drawImage(img, qrX, qrY, qrSize, qrSize);
                
                ctx.fillStyle = "#000000";
                ctx.font = `bold ${Math.max(7, Math.min(10, h * 0.3))}px Outfit, Arial, sans-serif`;
                ctx.textAlign = "center";
                ctx.textBaseline = "bottom";
                ctx.fillText(labelText, x + w / 2, topMargin + h - 4);
            }
        };
    }
    
    // Update preview page UI controls
    document.getElementById("preview-page-info").innerText = `Page ${previewCurrentPage} of ${previewTotalPages}`;
    document.getElementById("prev-preview-page").disabled = previewCurrentPage <= 1;
    document.getElementById("next-preview-page").disabled = previewCurrentPage >= previewTotalPages;
}

document.getElementById("prev-preview-page").addEventListener("click", () => {
    if (previewCurrentPage > 1) {
        previewCurrentPage--;
        drawPreviewPage();
    }
});

document.getElementById("next-preview-page").addEventListener("click", () => {
    if (previewCurrentPage < previewTotalPages) {
        previewCurrentPage++;
        drawPreviewPage();
    }
});

// Printing Sequential Batch to Local Print Agent
document.getElementById("btn-print-sequential").addEventListener("click", async () => {
    if (!printAgentConnected || currentGeneratedQRs.length === 0) return;
    
    const printer = "";
    const labelWidth = document.getElementById("seq-label-width").value;
    const labelHeight = document.getElementById("seq-label-height").value;
    const gap = document.getElementById("seq-gap").value;
    const topMargin = document.getElementById("seq-top-margin").value;
    
    const payload = {
        mode: "sequential",
        printer_name: printer,
        qrs: currentGeneratedQRs,
        label_width: labelWidth,
        label_height: labelHeight,
        gap: gap,
        top_margin: topMargin
    };
    
    const btn = document.getElementById("btn-print-sequential");
    btn.disabled = true;
    btn.innerText = "Spooling...";
    
    try {
        const response = await fetch("http://localhost:5001/print", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });
        const result = await response.json();
        
        if (result.success) {
            showGlobalAlert(`Batch of ${currentGeneratedQRs.length} spooled successfully to printer.`, "success");
        } else {
            showGlobalAlert(`Printing Error: ${result.message}`, "danger");
        }
    } catch (err) {
        showGlobalAlert("Failed to communicate with local Print Agent. Make sure print_agent.py is running.", "danger");
    } finally {
        btn.disabled = false;
        btn.innerText = "Print Current Batch";
    }
});

// ---------------------------------------------------------
// Tab: Customer PDF Operations Logic
// ---------------------------------------------------------
async function loadPdfGrid() {
    try {
        const response = await fetch("/api/pdf_list");
        const pdfs = await response.json();
        
        const tbody = document.querySelector("#pdf-grid tbody");
        tbody.innerHTML = "";
        
        if (pdfs.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="text-center">No PDFs imported yet.</td></tr>';
            return;
        }
        
        pdfs.forEach(pdf => {
            const tr = document.createElement("tr");
            tr.setAttribute("data-pdf-id", pdf.PDFID);
            tr.setAttribute("data-pdf-printed", pdf.Printed === 1 ? "Yes" : "No");
            tr.setAttribute("data-pdf-filename", pdf.FileName);
            
            tr.innerHTML = `
                <td>${pdf.PDFID}</td>
                <td>${pdf.ItemCode}</td>
                <td>${pdf.FileName}</td>
                <td>${pdf.TotalQty}</td>
                <td><span class="badge ${pdf.Printed === 1 ? 'badge-success' : 'badge-danger'}">${pdf.Printed === 1 ? 'Yes' : 'No'}</span></td>
            `;
            
            tr.addEventListener("click", () => selectPdfRow(tr));
            tbody.appendChild(tr);
        });
        
    } catch (err) {
        showGlobalAlert("Failed to load PDF records.", "danger");
    }
}

function selectPdfRow(rowElement) {
    const rows = document.querySelectorAll("#pdf-grid tbody tr");
    rows.forEach(r => r.classList.remove("selected"));
    
    rowElement.classList.add("selected");
    selectedPdfId = parseInt(rowElement.getAttribute("data-pdf-id"), 10);
    selectedPdfPrinted = rowElement.getAttribute("data-pdf-printed") === "Yes";
    selectedPdfFileName = rowElement.getAttribute("data-pdf-filename");
    
    // Enable print operations
    if (printAgentConnected) {
        document.getElementById("btn-print-pdf").disabled = false;
    }
    
    // Update button text depending on printed state
    const btn = document.getElementById("btn-print-pdf");
    if (selectedPdfPrinted) {
        btn.innerText = "Reprint Selected PDF";
        btn.className = "btn btn-danger btn-block";
    } else {
        btn.innerText = "Print Selected PDF";
        btn.className = "btn btn-success btn-block";
    }
}

// Custom page range enable/disable triggers
document.getElementById("pdf-range-check").addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    document.getElementById("pdf-from-page").disabled = !isChecked;
    document.getElementById("pdf-to-page").disabled = !isChecked;
});

// Handle PDF import form submission
document.getElementById("pdf-import-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    
    const form = document.getElementById("pdf-import-form");
    const formData = new FormData(form);
    const progressContainer = document.getElementById("import-progress-container");
    const progressBar = document.getElementById("import-progress-inner");
    const btn = document.getElementById("btn-import-pdf");
    
    progressContainer.classList.remove("hide");
    btn.disabled = true;
    
    // Animate mock progress bar
    let progress = 0;
    const progressInterval = setInterval(() => {
        if (progress < 90) {
            progress += 10;
            progressBar.style.width = `${progress}%`;
        }
    }, 200);
    
    try {
        const response = await fetch("/api/import_pdf", {
            method: "POST",
            body: formData
        });
        const result = await response.json();
        
        clearInterval(progressInterval);
        progressBar.style.width = "100%";
        
        setTimeout(() => {
            progressContainer.classList.add("hide");
            progressBar.style.width = "0%";
            btn.disabled = false;
            
            if (result.success) {
                showGlobalAlert(result.message, "success");
                form.reset();
                loadPdfGrid();
            } else {
                showGlobalAlert(result.message || "Import failed.", "danger");
            }
        }, 300);
        
    } catch (err) {
        clearInterval(progressInterval);
        progressContainer.classList.add("hide");
        btn.disabled = false;
        showGlobalAlert("Server network error during PDF import.", "danger");
    }
});

// Printing/Reprinting Selected PDF
document.getElementById("btn-print-pdf").addEventListener("click", () => {
    if (!selectedPdfId || !printAgentConnected) return;
    
    if (selectedPdfPrinted) {
        // Pop open supervisor validation modal
        openSupervisorModal();
    } else {
        // Direct print (first time)
        executePdfPrint(null);
    }
});

// Supervisor modal triggers
function openSupervisorModal() {
    document.getElementById("supervisor-modal").classList.remove("hide");
    document.getElementById("super-username").value = "";
    document.getElementById("super-password").value = "";
    document.getElementById("modal-error").classList.add("hide");
}

function closeSupervisorModal() {
    document.getElementById("supervisor-modal").classList.add("hide");
}

document.getElementById("btn-modal-cancel").addEventListener("click", closeSupervisorModal);

document.getElementById("btn-modal-authorize").addEventListener("click", async () => {
    const username = document.getElementById("super-username").value;
    const password = document.getElementById("super-password").value;
    const modalError = document.getElementById("modal-error");
    
    if (!username || !password) {
        modalError.innerText = "Please enter credentials.";
        modalError.classList.remove("hide");
        return;
    }
    
    try {
        const response = await fetch("/api/verify_supervisor", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ username, password })
        });
        const result = await response.json();
        
        if (result.success) {
            closeSupervisorModal();
            // Proceed to print with supervisor auth details
            executePdfPrint(result.supervisor_id);
        } else {
            modalError.innerText = result.message || "Invalid supervisor credentials.";
            modalError.classList.remove("hide");
        }
    } catch (err) {
        modalError.innerText = "Connection error.";
        modalError.classList.remove("hide");
    }
});

// Actual PDF download and transmission to local print agent
async function executePdfPrint(supervisorId) {
    const isRange = document.getElementById("pdf-range-check").checked;
    const fromPage = document.getElementById("pdf-from-page").value;
    const toPage = document.getElementById("pdf-to-page").value;
    const printer = "";
    
    let url = `/api/get_pdf_file/${selectedPdfId}`;
    if (isRange) {
        url += `?from_page=${fromPage}&to_page=${toPage}`;
    }
    
    const btn = document.getElementById("btn-print-pdf");
    btn.disabled = true;
    btn.innerText = "Downloading & Spooling...";
    
    try {
        // 1. Download file blob from central database server
        const fileResponse = await fetch(url);
        if (!fileResponse.ok) throw new Error("Failed to download PDF from server.");
        
        const blob = await fileResponse.blob();
        
        // 2. Convert Blob to Base64
        const reader = new FileReader();
        reader.onloadend = async () => {
            const base64Data = reader.result.split(',')[1];
            
            // 3. POST Base64 print data directly to local Print Agent
            const printPayload = {
                mode: "pdf",
                printer_name: printer,
                pdf_data: base64Data
            };
            
            try {
                const printResponse = await fetch("http://localhost:5001/print", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(printPayload)
                });
                const printResult = await printResponse.json();
                
                if (printResult.success) {
                    // 4. On successful local spooling, update printed status on central database
                    const updatePayload = {
                        pdf_id: selectedPdfId,
                        is_reprint: selectedPdfPrinted,
                        supervisor_id: supervisorId
                    };
                    
                    const updateResponse = await fetch("/api/update_print_status", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify(updatePayload)
                    });
                    const updateResult = await updateResponse.json();
                    
                    showGlobalAlert(`PDF spooled successfully. Database status updated.`, "success");
                    loadPdfGrid();
                } else {
                    showGlobalAlert(`Spooling Error: ${printResult.message}`, "danger");
                }
            } catch (err) {
                showGlobalAlert("Failed to print. Make sure local Print Agent is running.", "danger");
            } finally {
                btn.disabled = false;
                btn.innerText = selectedPdfPrinted ? "Reprint Selected PDF" : "Print Selected PDF";
            }
        };
        reader.readAsDataURL(blob);
        
    } catch (err) {
        showGlobalAlert(err.message || "Failed to process PDF file.", "danger");
        btn.disabled = false;
        btn.innerText = selectedPdfPrinted ? "Reprint Selected PDF" : "Print Selected PDF";
    }
}

// ---------------------------------------------------------
// Tab: Operational Reports Logic
// ---------------------------------------------------------
function initReports() {
    const btnLoad = document.getElementById("btn-load-report");
    const btnExport = document.getElementById("btn-export-csv");
    let currentReportRows = [];
    
    btnLoad.addEventListener("click", async () => {
        const reportType = parseInt(document.getElementById("report-select").value, 10);
        btnLoad.disabled = true;
        btnLoad.innerText = "Loading...";
        
        try {
            const response = await fetch("/api/reports", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ report_type: reportType })
            });
            const result = await response.json();
            
            if (result.success) {
                currentReportRows = result.data;
                renderReportTable(reportType, currentReportRows);
                btnExport.disabled = currentReportRows.length === 0;
            } else {
                showGlobalAlert(result.message || "Failed to run report.", "danger");
                btnExport.disabled = true;
            }
        } catch (err) {
            showGlobalAlert("Connection failure running SQL query.", "danger");
            btnExport.disabled = true;
        } finally {
            btnLoad.disabled = false;
            btnLoad.innerText = "Load Audit Report";
        }
    });
    
    btnExport.addEventListener("click", () => {
        if (currentReportRows.length === 0) return;
        
        // Grab headers from active grid
        const table = document.getElementById("reports-grid");
        const headers = Array.from(table.querySelectorAll("thead th")).map(th => th.innerText);
        
        // Construct CSV
        let csvContent = "\uFEFF"; // Add UTF-8 BOM for Excel formatting
        csvContent += headers.join(",") + "\n";
        
        const rows = table.querySelectorAll("tbody tr");
        rows.forEach(tr => {
            const cols = Array.from(tr.querySelectorAll("td")).map(td => {
                // Escape double quotes in CSV fields
                let txt = td.innerText.replace(/"/g, '""');
                return `"${txt}"`;
            });
            csvContent += cols.join(",") + "\n";
        });
        
        const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.setAttribute("download", `Report_Export_${new Date().toISOString().slice(0, 10)}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}

function renderReportTable(type, rows) {
    const table = document.getElementById("reports-grid");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody") || document.createElement("tbody");
    
    thead.innerHTML = "";
    tbody.innerHTML = "";
    
    if (rows.length === 0) {
        thead.innerHTML = "<tr><th>Report Data Grid</th></tr>";
        tbody.innerHTML = '<tr><td class="text-center">No rows returned for this query.</td></tr>';
        table.appendChild(tbody);
        return;
    }
    
    const trHead = document.createElement("tr");
    
    // Dynamic headers depending on select type
    let keys = [];
    if (type === 0) {
        trHead.innerHTML = "<th>Row</th><th>Reprinted On</th><th>QR Value / PDF Filename</th><th>Authorized By</th><th>Reason</th>";
        keys = ["ReprintedOn", "QRValue", "AuthorizedBy", "Reason"];
    } else if (type === 1) {
        trHead.innerHTML = "<th>Row</th><th>Model</th><th>QR Code Value</th><th>Registry Date</th><th>Packaging Scan Date</th><th>Current Status</th>";
        keys = ["ItemCode", "QRValue", "GeneratedOn", "ScanDate", "Status"];
    } else if (type === 2) {
        trHead.innerHTML = "<th>Row</th><th>PDF ID</th><th>Model</th><th>Customer PDF Filename</th><th>Qty</th><th>Printed?</th><th>Printed On</th><th>Reprints</th>";
        keys = ["PDFID", "ItemCode", "FileName", "TotalQty", "Printed", "PrintedOn", "ReprintCount"];
    } else if (type === 3) {
        trHead.innerHTML = "<th>Row</th><th>Date</th><th>Model Code</th><th>Total QR Codes Generated</th><th>Total QR Codes Printed</th>";
        keys = ["ProdDate", "ItemCode", "TotalGenerated", "TotalPrinted"];
    }
    thead.appendChild(trHead);
    
    // Draw rows
    rows.forEach((row, i) => {
        const tr = document.createElement("tr");
        const tdRow = `<td>${i + 1}</td>`;
        
        const tdContent = keys.map(k => {
            let val = row[k];
            if (val === null || val === undefined) val = "N/A";
            
            // Format true/false or binary
            if (k === "Printed" && (val === 1 || val === "1")) val = "Yes";
            if (k === "Printed" && (val === 0 || val === "0")) val = "No";
            
            return `<td>${val}</td>`;
        }).join("");
        
        tr.innerHTML = tdRow + tdContent;
        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
}

// ---------------------------------------------------------
// Tab: User Management Logic (Admin Only)
// ---------------------------------------------------------
function initUserManagement() {
    const grid = document.getElementById("users-grid");
    if (!grid) return; // Exit if user is not admin
    
    const userField = document.getElementById("user-id-input");
    const passField = document.getElementById("user-password-input");
    const roleField = document.getElementById("user-role-select");
    
    const btnAdd = document.getElementById("btn-user-add");
    const btnUpdate = document.getElementById("btn-user-update");
    const btnDelete = document.getElementById("btn-user-delete");
    const btnClear = document.getElementById("btn-user-clear");
    
    async function loadUserGrid() {
        try {
            const response = await fetch("/api/users");
            const users = await response.json();
            
            const tbody = grid.querySelector("tbody");
            tbody.innerHTML = "";
            
            users.forEach(u => {
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td>${u.UserID}</td>
                    <td>${u.PasswordHash}</td>
                    <td>${u.UserRole}</td>
                `;
                tr.addEventListener("click", () => {
                    // Populate inputs
                    userField.value = u.UserID;
                    passField.value = u.PasswordHash;
                    roleField.value = u.UserRole;
                    
                    // Lock User ID (primary key modification safety)
                    userField.disabled = true;
                    
                    // Set active row highlight
                    grid.querySelectorAll("tbody tr").forEach(r => r.classList.remove("selected"));
                    tr.classList.add("selected");
                    
                    // Toggle CRUD button enabling states (safeguard UX rule)
                    btnAdd.disabled = true;
                    btnUpdate.disabled = false;
                    btnDelete.disabled = false;
                });
                tbody.appendChild(tr);
            });
        } catch (err) {
            showGlobalAlert("Failed to load user master database.", "danger");
        }
    }
    
    function clearFields() {
        userField.value = "";
        passField.value = "";
        roleField.value = "Operator";
        userField.disabled = false;
        
        grid.querySelectorAll("tbody tr").forEach(r => r.classList.remove("selected"));
        
        btnAdd.disabled = false;
        btnUpdate.disabled = true;
        btnDelete.disabled = true;
    }
    
    btnClear.addEventListener("click", clearFields);
    
    // Add user
    btnAdd.addEventListener("click", async () => {
        const username = userField.value.trim();
        const password = passField.value.trim();
        const role = roleField.value;
        
        if (!username || !password || !role) {
            showGlobalAlert("Please fill in all user credentials details.", "danger");
            return;
        }
        
        try {
            const response = await fetch("/api/users", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ username, password, role })
            });
            const result = await response.json();
            
            if (result.success) {
                showGlobalAlert(result.message, "success");
                clearFields();
                loadUserGrid();
            } else {
                showGlobalAlert(result.message || "Failed to create user.", "danger");
            }
        } catch (err) {
            showGlobalAlert("Network connection error.", "danger");
        }
    });
    
    // Update user
    btnUpdate.addEventListener("click", async () => {
        const username = userField.value.trim();
        const password = passField.value.trim();
        const role = roleField.value;
        
        if (!username || !password || !role) {
            showGlobalAlert("Missing fields.", "danger");
            return;
        }
        
        try {
            const response = await fetch(`/api/users/${encodeURIComponent(username)}`, {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ password, role })
            });
            const result = await response.json();
            
            if (result.success) {
                showGlobalAlert(result.message, "success");
                clearFields();
                loadUserGrid();
            } else {
                showGlobalAlert(result.message || "Failed to update user.", "danger");
            }
        } catch (err) {
            showGlobalAlert("Network connection error.", "danger");
        }
    });
    
    // Delete user
    btnDelete.addEventListener("click", async () => {
        const username = userField.value.trim();
        if (!username) return;
        
        if (!confirm(`Are you sure you want to delete user account '${username}'?`)) {
            return;
        }
        
        try {
            const response = await fetch(`/api/users/${encodeURIComponent(username)}`, {
                method: "DELETE"
            });
            const result = await response.json();
            
            if (result.success) {
                showGlobalAlert(result.message, "success");
                clearFields();
                loadUserGrid();
            } else {
                showGlobalAlert(result.message || "Failed to delete user.", "danger");
            }
        } catch (err) {
            showGlobalAlert("Network connection error.", "danger");
        }
    });
    
    // Initial user grid loading
    loadUserGrid();
}
