let componentsData = {
    cpu: [],
    motherboard: [],
    ram: [],
    storageType: [
        { id: 'sata', name: 'SSD SATA' },
        { id: 'nvme', name: 'SSD M.2 NVMe' },
        { id: 'hdd', name: 'Disco Duro HDD' }
    ],
    storage: {
        sata: [],
        nvme: [],
        hdd: []
    },
    case: [],
    psu: [],
    gpu: [],
    cooling: []
};

async function loadExcelData() {
    try {
        const response = await fetch('Lista para Armador pc.xlsx');
        if (!response.ok) throw new Error("No se pudo cargar el archivo Excel");
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        let currentCategory = '';
        
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            
            if (!row || row.length === 0 || (row.length === 1 && !row[0])) continue;
            
            if (row.length === 1 || (row.length >= 2 && !row[1])) {
                currentCategory = String(row[0]).trim().toLowerCase();
                continue;
            }
            
            const name = String(row[0]).trim();
            const price = parseFloat(row[1]) || 0;
            const id = 'item_' + i;
            const product = { id, name, price };
            
            if (currentCategory.includes('procesador')) {
                let socket = 'AM4';
                if (currentCategory.includes('am5') || name.includes('AM5') || name.includes('7600') || name.includes('7700') || name.includes('7900') || name.includes('8500') || name.includes('8700')) socket = 'AM5';
                else if (currentCategory.includes('am4') || name.includes('AM4')) socket = 'AM4';
                else if (currentCategory.includes('1700') || name.includes('1700') || name.includes('12100') || name.includes('12400') || name.includes('13400') || name.includes('14700') || name.includes('12700')) socket = 'LGA1700';
                else if (currentCategory.includes('1200') || name.includes('1200') || name.includes('10500')) socket = 'LGA1200';
                else if (currentCategory.includes('1851') || name.includes('1851')) socket = 'LGA1851';
                
                product.socket = socket;
                componentsData.cpu.push(product);
            }
            else if (currentCategory.includes('mother')) {
                let socket = 'AM4';
                if (currentCategory.includes('am5') || name.toUpperCase().includes('B650') || name.toUpperCase().includes('X670') || name.toUpperCase().includes('A620') || name.toUpperCase().includes('AM5')) socket = 'AM5';
                else if (currentCategory.includes('am4') || name.toUpperCase().includes('B550') || name.toUpperCase().includes('A520') || name.toUpperCase().includes('X570') || name.toUpperCase().includes('AM4')) socket = 'AM4';
                else if (currentCategory.includes('1700') || name.toUpperCase().includes('H610') || name.toUpperCase().includes('B760') || name.toUpperCase().includes('Z790') || name.toUpperCase().includes('1700')) socket = 'LGA1700';
                else if (currentCategory.includes('1200') || name.toUpperCase().includes('H510') || name.toUpperCase().includes('B560') || name.toUpperCase().includes('1200')) socket = 'LGA1200';
                else if (currentCategory.includes('1851')) socket = 'LGA1851';
                
                product.socket = socket;
                componentsData.motherboard.push(product);
            }
            else if (currentCategory.includes('memoria') && !currentCategory.includes('video')) {
                componentsData.ram.push(product);
            }
            else if (currentCategory.includes('disco') || currentCategory.includes('ssd') || currentCategory.includes('almacenamiento')) {
                if (currentCategory.includes('nvme') || name.toLowerCase().includes('nvme') || name.toLowerCase().includes('m.2')) {
                    componentsData.storage.nvme.push(product);
                } else if (currentCategory.includes('hdd') || name.toLowerCase().includes('hdd') || currentCategory.includes('duro')) {
                    componentsData.storage.hdd.push(product);
                } else {
                    componentsData.storage.sata.push(product);
                }
            }
            else if (currentCategory.includes('gabinete')) {
                componentsData.case.push(product);
            }
            else if (currentCategory.includes('fuente') || currentCategory.includes('psu')) {
                componentsData.psu.push(product);
            }
            else if (currentCategory.includes('video') || currentCategory.includes('vga') || currentCategory.includes('gpu') || currentCategory.includes('placa')) {
                componentsData.gpu.push(product);
            }
            else if (currentCategory.includes('cooler') || currentCategory.includes('refrigeraci') || currentCategory.includes('water') || currentCategory.includes('coler')) {
                componentsData.cooling.push(product);
            }
        }
        
        if (componentsData.cooling.length === 0) {
            componentsData.cooling.push({ id: 'co_stock', name: 'Cooler Stock (Incluido con CPU)', price: 0 });
        }
        if (componentsData.gpu.length === 0) {
            componentsData.gpu.push({ id: 'gpu_int', name: 'Sin placa (Integrados)', price: 0 });
        }
        if (componentsData.psu.length === 0) {
            componentsData.psu.push({ id: 'psu_gen', name: 'Fuente', price: 0 });
        }
        
    } catch (error) {
        console.error("Error loading Excel:", error);
        alert("Hubo un error cargando la base de datos de precios desde Excel.");
    }
}
// DOM Elements
const platformBtns = document.querySelectorAll('.platform-btn');
const cpuSelect = document.getElementById('cpu');
const mbSelect = document.getElementById('motherboard');
const ramSelect = document.getElementById('ram');
const ramQty = document.getElementById('ram-qty');
const storageList = document.getElementById('storage-list');
const addStorageBtn = document.getElementById('add-storage-btn');
const caseSelect = document.getElementById('case');
const psuSelect = document.getElementById('psu');
const gpuSelect = document.getElementById('gpu');
const coolingSelect = document.getElementById('cooling');
const totalPriceEl = document.getElementById('total-price');
const copyBtn = document.getElementById('copy-config-btn');
const toast = document.getElementById('toast');
const buyMultipleCheckbox = document.getElementById('buy-multiple');
const whatsappInfo = document.getElementById('whatsapp-info');
const priceLabel = document.querySelector('.price-label');
const whatsappBtn = document.getElementById('whatsapp-btn');

// Theme Elements
const html = document.documentElement;
const btnLight = document.getElementById('btn-light-theme');
const btnDark = document.getElementById('btn-dark-theme');

// State
let selectedPlatform = null;
let storageCounter = 0;
let isRestoring = false;

function saveConfiguratorState() {
    if (isRestoring) return;
    
    const storageRows = [];
    document.querySelectorAll('.storage-row').forEach(row => {
        const typeSelect = row.querySelector('.storage-type-select');
        const itemSelect = row.querySelector('.storage-item-select');
        if (typeSelect && itemSelect) {
            storageRows.push({
                type: typeSelect.value,
                item: itemSelect.value
            });
        }
    });

    const state = {
        selectedPlatform: selectedPlatform,
        cpu: cpuSelect ? cpuSelect.value : '',
        motherboard: mbSelect ? mbSelect.value : '',
        ram: ramSelect ? ramSelect.value : '',
        ramQty: ramQty ? ramQty.value : '1',
        storageRows: storageRows,
        case: caseSelect ? caseSelect.value : '',
        psu: psuSelect ? psuSelect.value : '',
        gpu: gpuSelect ? gpuSelect.value : '',
        cooling: coolingSelect ? coolingSelect.value : '',
        buyMultiple: buyMultipleCheckbox ? buyMultipleCheckbox.checked : false
    };

    localStorage.setItem('configurator_saved_state', JSON.stringify(state));
}

function restoreConfiguratorState() {
    const savedStateStr = localStorage.getItem('configurator_saved_state');
    if (!savedStateStr) return;
    
    isRestoring = true;
    try {
        const state = JSON.parse(savedStateStr);
        
        if (state.selectedPlatform) {
            const btn = Array.from(platformBtns).find(b => b.getAttribute('data-socket') === state.selectedPlatform);
            if (btn) {
                btn.click();
            }
        }
        
        if (state.cpu && cpuSelect) {
            cpuSelect.value = state.cpu;
            cpuSelect.dispatchEvent(new Event('change'));
        }
        
        if (state.motherboard && mbSelect) {
            mbSelect.value = state.motherboard;
            mbSelect.dispatchEvent(new Event('change'));
        }
        
        if (state.ram && ramSelect) {
            ramSelect.value = state.ram;
            ramSelect.dispatchEvent(new Event('change'));
        }
        if (state.ramQty && ramQty) {
            ramQty.value = state.ramQty;
            ramQty.dispatchEvent(new Event('change'));
        }
        
        if (state.storageRows && state.storageRows.length > 0) {
            storageList.innerHTML = '';
            storageCounter = 0;
            
            state.storageRows.forEach(savedRow => {
                addStorageRow();
                const rows = storageList.querySelectorAll('.storage-row');
                const lastRow = rows[rows.length - 1];
                if (lastRow) {
                    const typeSelect = lastRow.querySelector('.storage-type-select');
                    const itemSelect = lastRow.querySelector('.storage-item-select');
                    if (typeSelect && itemSelect) {
                        typeSelect.value = savedRow.type;
                        typeSelect.dispatchEvent(new Event('change'));
                        itemSelect.value = savedRow.item;
                        itemSelect.dispatchEvent(new Event('change'));
                    }
                }
            });
        }
        
        if (state.case && caseSelect) {
            caseSelect.value = state.case;
            caseSelect.dispatchEvent(new Event('change'));
        }
        if (state.psu && psuSelect) {
            psuSelect.value = state.psu;
            psuSelect.dispatchEvent(new Event('change'));
        }
        if (state.gpu && gpuSelect) {
            gpuSelect.value = state.gpu;
            gpuSelect.dispatchEvent(new Event('change'));
        }
        if (state.cooling && coolingSelect) {
            coolingSelect.value = state.cooling;
            coolingSelect.dispatchEvent(new Event('change'));
        }
        
        if (state.buyMultiple && buyMultipleCheckbox) {
            buyMultipleCheckbox.checked = state.buyMultiple;
            buyMultipleCheckbox.dispatchEvent(new Event('change'));
        }
        
    } catch (e) {
        console.error("Error restoring configuration state:", e);
    } finally {
        isRestoring = false;
        updatePrice();
    }
}

function initCountdown() {
    const COUNTDOWN_KEY = 'promo_target_time';
    const THREE_DAYS_MS = 3 * 24 * 60 * 60 * 1000;
    
    let targetTime = localStorage.getItem(COUNTDOWN_KEY);
    const now = Date.now();
    
    if (!targetTime) {
        targetTime = now + THREE_DAYS_MS;
        localStorage.setItem(COUNTDOWN_KEY, targetTime);
    } else {
        targetTime = parseInt(targetTime, 10);
        if (isNaN(targetTime) || now >= targetTime) {
            targetTime = now + THREE_DAYS_MS;
            localStorage.setItem(COUNTDOWN_KEY, targetTime);
        }
    }
    
    const daysEl = document.getElementById('cd-days');
    const hoursEl = document.getElementById('cd-hours');
    const minutesEl = document.getElementById('cd-minutes');
    const secondsEl = document.getElementById('cd-seconds');
    const footerEl = document.getElementById('cd-footer');
    
    function updateCountdown() {
        const currentTime = Date.now();
        const timeLeft = targetTime - currentTime;
        
        if (timeLeft <= 0) {
            if (daysEl) daysEl.textContent = '00';
            if (hoursEl) hoursEl.textContent = '00';
            if (minutesEl) minutesEl.textContent = '00';
            if (secondsEl) secondsEl.textContent = '00';
            if (footerEl) footerEl.textContent = '¡Promoción finalizada!';
            clearInterval(timerInterval);
            return;
        }
        
        const days = Math.floor(timeLeft / (1000 * 60 * 60 * 24));
        const hours = Math.floor((timeLeft % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        const minutes = Math.floor((timeLeft % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((timeLeft % (1000 * 60)) / 1000);
        
        if (daysEl) daysEl.textContent = String(days).padStart(2, '0');
        if (hoursEl) hoursEl.textContent = String(hours).padStart(2, '0');
        if (minutesEl) minutesEl.textContent = String(minutes).padStart(2, '0');
        if (secondsEl) secondsEl.textContent = String(seconds).padStart(2, '0');
        
        if (footerEl) {
            if (days > 0) {
                footerEl.textContent = `Finaliza en ${days} ${days === 1 ? 'día' : 'días'}`;
            } else {
                footerEl.textContent = 'Finaliza hoy';
            }
        }
    }
    
    updateCountdown();
    const timerInterval = setInterval(updateCountdown, 1000);
}

// Initialize
async function init() {
    await loadExcelData();
    
    initTheme();
    initCountdown();
    populateSelect(caseSelect, componentsData.case);
    populateSelect(psuSelect, componentsData.psu);
    populateSelect(gpuSelect, componentsData.gpu);
    populateSelect(coolingSelect, componentsData.cooling);
    
    // Only restore state or set default storage row if no query parameters exist
    const hasQueryParams = window.location.search.length > 0;
    if (hasQueryParams) {
        addStorageRow(); // Initial storage row
        loadFromQueryParams();
    } else {
        const savedState = localStorage.getItem('configurator_saved_state');
        if (savedState) {
            restoreConfiguratorState();
        } else {
            addStorageRow(); // Initial default storage row
        }
    }
    
    // Listeners
    platformBtns.forEach(btn => btn.addEventListener('click', handlePlatformClick));
    mbSelect.addEventListener('change', handleMotherboardChange);
    addStorageBtn.addEventListener('click', addStorageRow);
    copyBtn.addEventListener('click', copyConfig);
    if (whatsappBtn) {
        whatsappBtn.addEventListener('click', handleWhatsAppClick);
    }
    
    buyMultipleCheckbox.addEventListener('change', () => {
        if (buyMultipleCheckbox.checked) {
            whatsappInfo.classList.add('show');
        } else {
            whatsappInfo.classList.remove('show');
        }
        updatePrice();
    });
    
    // Price update listeners
    document.querySelector('.configurator-main').addEventListener('change', updatePrice);
}

// Theme
function initTheme() {
    const savedTheme = localStorage.getItem('theme') || 'dark';
    html.setAttribute('data-theme', savedTheme);
    if (savedTheme === 'dark') {
        if (btnDark) btnDark.classList.add('active');
        if (btnLight) btnLight.classList.remove('active');
    } else {
        if (btnLight) btnLight.classList.add('active');
        if (btnDark) btnDark.classList.remove('active');
    }

    if (btnLight) {
        btnLight.addEventListener('click', () => {
            html.setAttribute('data-theme', 'light');
            localStorage.setItem('theme', 'light');
            btnLight.classList.add('active');
            btnDark.classList.remove('active');
        });
    }
    
    if (btnDark) {
        btnDark.addEventListener('click', () => {
            html.setAttribute('data-theme', 'dark');
            localStorage.setItem('theme', 'dark');
            btnDark.classList.add('active');
            btnLight.classList.remove('active');
        });
    }
}

// Platform / Socket
function handlePlatformClick(e) {
    platformBtns.forEach(btn => btn.classList.remove('active'));
    e.target.classList.add('active');
    selectedPlatform = e.target.getAttribute('data-socket');
    
    // Filter CPUs and MBs
    const filteredCPUs = componentsData.cpu.filter(c => c.socket === selectedPlatform);
    const filteredMBs = componentsData.motherboard.filter(m => m.socket === selectedPlatform);
    
    populateSelect(cpuSelect, filteredCPUs, '-- Seleccione Procesador --');
    cpuSelect.disabled = false;
    
    populateSelect(mbSelect, filteredMBs, '-- Seleccione Motherboard --');
    mbSelect.disabled = false;
    
    // Reset RAM
    ramSelect.innerHTML = '<option value="">-- Seleccione Mother primero --</option>';
    ramSelect.disabled = true;
    ramQty.disabled = true;
    
    updatePrice();
}

// Motherboard
function handleMotherboardChange() {
    if (mbSelect.value) {
        populateSelect(ramSelect, componentsData.ram, '-- Seleccione Memoria RAM --');
        ramSelect.disabled = false;
        ramQty.disabled = false;
    } else {
        ramSelect.innerHTML = '<option value="">-- Seleccione Mother primero --</option>';
        ramSelect.disabled = true;
        ramQty.disabled = true;
    }
}

// Utility to populate select
function populateSelect(selectEl, dataList, defaultText = '-- Seleccione --') {
    selectEl.innerHTML = `<option value="">${defaultText}</option>`;
    dataList.forEach(item => {
        const option = document.createElement('option');
        option.value = item.id;
        option.textContent = `${item.name} - $${item.price.toLocaleString('es-AR')}`;
        option.dataset.price = item.price;
        option.dataset.name = item.name;
        selectEl.appendChild(option);
    });
}

// Storage Dynamic Rows
function addStorageRow() {
    storageCounter++;
    const rowId = `storage-row-${storageCounter}`;
    
    const row = document.createElement('div');
    row.className = 'storage-row row-group';
    row.id = rowId;
    
    // Type Select
    const typeWrapper = document.createElement('div');
    typeWrapper.className = 'storage-type';
    typeWrapper.innerHTML = `
        <div class="select-wrapper">
            <select class="custom-select storage-type-select">
                ${componentsData.storageType.map(t => `<option value="${t.id}">${t.name}</option>`).join('')}
            </select>
            <i class="ph-bold ph-caret-down select-icon"></i>
        </div>
    `;
    
    // Item Select
    const itemWrapper = document.createElement('div');
    itemWrapper.className = 'storage-item';
    const itemSelectHTML = `
        <div class="select-wrapper">
            <select class="custom-select storage-item-select">
                <option value="">-- Seleccione --</option>
                ${componentsData.storage['sata'].map(item => `<option value="${item.id}" data-price="${item.price}" data-name="${item.name}">${item.name} - $${item.price.toLocaleString('es-AR')}</option>`).join('')}
            </select>
            <i class="ph-bold ph-caret-down select-icon"></i>
        </div>
    `;
    itemWrapper.innerHTML = itemSelectHTML;
    
    // Remove Button
    const removeBtn = document.createElement('button');
    removeBtn.className = 'btn-remove';
    removeBtn.innerHTML = '<i class="ph-bold ph-x"></i>';
    if (storageCounter === 1) {
        removeBtn.style.visibility = 'hidden'; // Don't remove the first one easily, or allow it
    }
    removeBtn.addEventListener('click', () => {
        row.remove();
        updatePrice();
    });
    
    // Listen to type change
    const typeSelect = typeWrapper.querySelector('.storage-type-select');
    const itemSelect = itemWrapper.querySelector('.storage-item-select');
    
    typeSelect.addEventListener('change', (e) => {
        const selectedType = e.target.value;
        const items = componentsData.storage[selectedType];
        
        itemSelect.innerHTML = '<option value="">-- Seleccione --</option>';
        items.forEach(item => {
            const opt = document.createElement('option');
            opt.value = item.id;
            opt.dataset.price = item.price;
            opt.dataset.name = item.name;
            opt.textContent = `${item.name} - $${item.price.toLocaleString('es-AR')}`;
            itemSelect.appendChild(opt);
        });
        updatePrice();
    });
    
    row.appendChild(typeWrapper);
    row.appendChild(itemWrapper);
    row.appendChild(removeBtn);
    
    storageList.appendChild(row);
}

// Calculate Price and Update Summary
function updatePrice() {
    let total = 0;
    const summaryContent = document.getElementById('summary-content');
    summaryContent.innerHTML = ''; // clear current
    
    let hasItems = false;
    
    // Helper to get price and name from select
    const addItemToSummary = (selectEl, label, multiplier = 1) => {
        if (selectEl && selectEl.value) {
            hasItems = true;
            const opt = selectEl.options[selectEl.selectedIndex];
            const name = opt.dataset.name;
            const price = parseFloat(opt.dataset.price) || 0;
            const lineTotal = price * multiplier;
            
            total += lineTotal;
            
            const qtyStr = multiplier > 1 ? ` (x${multiplier})` : '';
            const itemDiv = document.createElement('div');
            itemDiv.className = 'summary-item';
            itemDiv.innerHTML = `
                <span class="summary-item-name">${label}: ${name}${qtyStr}</span>
                <span class="summary-item-price">$${lineTotal.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</span>
            `;
            summaryContent.appendChild(itemDiv);
        }
    };
    
    addItemToSummary(cpuSelect, 'CPU');
    addItemToSummary(mbSelect, 'Motherboard');
    
    if (ramSelect.value) {
        const qty = parseInt(ramQty.value) || 1;
        addItemToSummary(ramSelect, 'RAM', qty);
    }
    
    // Storage
    document.querySelectorAll('.storage-item-select').forEach((sel, index) => {
        if (sel.value) {
            addItemToSummary(sel, `Almacenamiento ${index + 1}`);
        }
    });
    
    addItemToSummary(caseSelect, 'Gabinete');
    addItemToSummary(psuSelect, 'Fuente');
    addItemToSummary(gpuSelect, 'Video');
    addItemToSummary(coolingSelect, 'Refrigeración');
    
    if (!hasItems) {
        summaryContent.innerHTML = '<p class="empty-summary">Aún no has seleccionado componentes.</p>';
    }
    
    const isBulk = buyMultipleCheckbox.checked;
    const finalTotal = isBulk ? total * 3 : total;
    
    if (isBulk) {
        priceLabel.innerHTML = 'TOTAL APROX (SIN DESC.) <i class="ph-bold ph-whatsapp" style="color: #25D366; font-size: 16px; vertical-align: middle;"></i>';
        totalPriceEl.textContent = `$ ${finalTotal.toLocaleString('es-AR', {minimumFractionDigits: 0, maximumFractionDigits: 0})} (x3)`;
    } else {
        priceLabel.textContent = 'TOTAL IVA INCLUÍDO';
        totalPriceEl.textContent = `$ ${finalTotal.toLocaleString('es-AR', {minimumFractionDigits: 0, maximumFractionDigits: 0})}`;
    }
    
    saveConfiguratorState();
}

function generateConfigText(isBulk) {
    let configText = isBulk ? "--- MI CONFIGURACIÓN DE PC (X3 MAYORISTA) ---\n\n" : "--- MI CONFIGURACIÓN DE PC ---\n\n";
    let total = 0;
    
    const appendItem = (label, selectEl, multiplier = 1) => {
        if (selectEl && selectEl.value) {
            const opt = selectEl.options[selectEl.selectedIndex];
            const name = opt.dataset.name;
            const price = parseFloat(opt.dataset.price);
            const lineTotal = price * multiplier;
            
            let qtyStr = multiplier > 1 ? ` (x${multiplier})` : '';
            configText += `${label}: ${name}${qtyStr} - $${lineTotal.toLocaleString('es-AR')}\n`;
            total += lineTotal;
        }
    };
    
    if (selectedPlatform) {
        configText += `Plataforma: ${selectedPlatform}\n`;
    }
    
    appendItem('Procesador', cpuSelect);
    appendItem('Motherboard', mbSelect);
    
    if (ramSelect.value) {
        appendItem('Memoria RAM', ramSelect, parseInt(ramQty.value));
    }
    
    const storageSelects = document.querySelectorAll('.storage-item-select');
    storageSelects.forEach((sel, index) => {
        if (sel.value) {
            appendItem(`Almacenamiento ${index + 1}`, sel);
        }
    });
    
    appendItem('Gabinete', caseSelect);
    appendItem('Fuente', psuSelect);
    appendItem('Placa de Video', gpuSelect);
    appendItem('Refrigeración', coolingSelect);
    
    if (isBulk) {
        configText += `\nCANTIDAD SELECCIONADA: 3 UNIDADES\n`;
        configText += `TOTAL APROXIMADO (Sin descuento): $${(total * 3).toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}\n`;
        configText += `\n¡Hola! Quiero consultar por el descuento mayorista para esta configuración.`;
    } else {
        configText += `\nTOTAL IVA INCLUIDO: $${total.toLocaleString('es-AR', {minimumFractionDigits: 2, maximumFractionDigits: 2})}\n`;
    }
    
    return configText;
}

// Copy Configuration
function copyConfig() {
    const isBulk = buyMultipleCheckbox && buyMultipleCheckbox.checked;
    const configText = generateConfigText(isBulk);
    
    navigator.clipboard.writeText(configText).then(() => {
        showToast();
    }).catch(err => {
        console.error("Error al copiar", err);
        alert("No se pudo copiar al portapapeles.");
    });
}

function handleWhatsAppClick() {
    const configText = generateConfigText(true);
    const encodedText = encodeURIComponent(configText);
    const waUrl = `https://wa.me/5491131184780?text=${encodedText}`;
    window.open(waUrl, '_blank');
}

function showToast() {
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Start

function loadFromQueryParams() {
    const params = new URLSearchParams(window.location.search);
    const socket = params.get('socket');
    const cpu = params.get('cpu');
    const motherboard = params.get('motherboard');
    const ram = params.get('ram');
    const ramQtyVal = params.get('ramQty');
    const storageType = params.get('storageType');
    const storage = params.get('storage');
    const caseId = params.get('case');
    
    if (socket) {
        const btn = Array.from(platformBtns).find(b => b.getAttribute('data-socket') === socket);
        if (btn) {
            btn.click();
        }
    }
    
    if (cpu && cpuSelect) {
        cpuSelect.value = cpu;
        cpuSelect.dispatchEvent(new Event('change'));
    }
    
    if (motherboard && mbSelect) {
        mbSelect.value = motherboard;
        mbSelect.dispatchEvent(new Event('change'));
    }
    
    if (ram && ramSelect) {
        ramSelect.value = ram;
        ramSelect.dispatchEvent(new Event('change'));
    }
    
    if (ramQtyVal && ramQty) {
        ramQty.value = ramQtyVal;
        ramQty.dispatchEvent(new Event('change'));
    }
    
    if (storageType && storage) {
        const row = storageList.querySelector('.storage-row');
        if (row) {
            const typeSelect = row.querySelector('.storage-type-select');
            const itemSelect = row.querySelector('.storage-item-select');
            if (typeSelect && itemSelect) {
                typeSelect.value = storageType;
                typeSelect.dispatchEvent(new Event('change'));
                itemSelect.value = storage;
                itemSelect.dispatchEvent(new Event('change'));
            }
        }
    }
    
    if (caseId && caseSelect) {
        caseSelect.value = caseId;
        caseSelect.dispatchEvent(new Event('change'));
    }
    
    updatePrice();
}

// --- Mobile Experience Enhancements ---

// Haptic Feedback Utility
function triggerHaptic() {
    if (navigator.vibrate) {
        navigator.vibrate(10); // 10ms subtle vibration
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // Add haptic feedback to selects and buttons
    const interactiveElements = document.querySelectorAll('select, button, .platform-btn, input[type="checkbox"]');
    interactiveElements.forEach(el => {
        el.addEventListener('change', triggerHaptic);
        el.addEventListener('click', (e) => {
            if (el.tagName === 'BUTTON' || el.classList.contains('platform-btn')) {
                triggerHaptic();
            }
        });
    });

    // Custom Share Modal Logic
    const shareModal = document.getElementById('share-modal');
    const shareModalClose = document.getElementById('share-modal-close');
    const shareBtn = document.getElementById('share-btn');
    const shareWp = document.getElementById('share-whatsapp-btn');
    const shareEmail = document.getElementById('share-email-btn');
    const shareCopy = document.getElementById('share-copy-btn');
    
    let currentShareText = "";

    function openShareModal(text) {
        currentShareText = text;
        if (shareModal) shareModal.classList.add('active');
        triggerHaptic();
    }

    if (shareModalClose) {
        shareModalClose.addEventListener('click', () => {
            shareModal.classList.remove('active');
        });
    }

    if (shareBtn) {
        shareBtn.style.display = 'flex';
        shareBtn.addEventListener('click', () => {
            const isBulk = document.getElementById('buy-multiple') ? document.getElementById('buy-multiple').checked : false;
            // Assuming generateConfigText is accessible
            const configText = generateConfigText(isBulk);
            openShareModal(configText);
        });
    }

    if (shareWp) {
        shareWp.addEventListener('click', () => {
            const waUrl = `https://wa.me/?text=${encodeURIComponent(currentShareText)}`;
            window.open(waUrl, '_blank');
            triggerHaptic();
            shareModal.classList.remove('active');
        });
    }

    if (shareEmail) {
        shareEmail.addEventListener('click', () => {
            const mailUrl = `mailto:?subject=${encodeURIComponent("Presupuesto de PC a Medida")}&body=${encodeURIComponent(currentShareText)}`;
            window.open(mailUrl, '_blank');
            triggerHaptic();
            shareModal.classList.remove('active');
        });
    }

    if (shareCopy) {
        shareCopy.addEventListener('click', () => {
            navigator.clipboard.writeText(currentShareText).then(() => {
                const toastEl = document.getElementById('toast');
                if (toastEl) {
                    toastEl.classList.add('show');
                    setTimeout(() => toastEl.classList.remove('show'), 3000);
                }
                shareModal.classList.remove('active');
            });
            triggerHaptic();
        });
    }
});

document.addEventListener('DOMContentLoaded', init);
