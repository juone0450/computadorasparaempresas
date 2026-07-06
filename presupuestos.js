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

// Define predefined builds and their matching rules
const predefinedPcs = {
    amd: [
        {
            title: "PC AMD Hogar / Oficina",
            desc: "Ideal para tareas cotidianas, navegación web y trabajo administrativo de oficina.",
            specs: {
                cpu: "3200G",
                motherboard: "A520M",
                ram: "DDR4 8GB 3200",
                ramQty: 1,
                storage: "SSD 240GB",
                storageType: "sata",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        },
        {
            title: "PC AMD Gaming Entrada",
            desc: "Excelente rendimiento para juegos competitivos (Valorant, CS:GO, LoL) en resolución Full HD.",
            specs: {
                cpu: "5600GT",
                motherboard: "A520M",
                ram: "DDR4 16GB 3200",
                ramQty: 1,
                storage: "SSD M.2 NVME 512GB",
                storageType: "nvme",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        },
        {
            title: "PC AMD Alto Rendimiento",
            desc: "Poder de sobra para streaming, edición de video y gaming exigente.",
            specs: {
                cpu: "5700G",
                motherboard: "B550M",
                ram: "DDR4 16GB 3200",
                ramQty: 2,
                storage: "M.2 NVME 512GB",
                storageType: "nvme",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        }
    ],
    intel: [
        {
            title: "PC Intel Oficina Eficiente",
            desc: "Estable, rápida y eficiente. Diseñada para contabilidad, navegación y uso hogareño.",
            specs: {
                cpu: "12100",
                motherboard: "H610M",
                ram: "DDR4 8GB 3200",
                ramQty: 1,
                storage: "SSD 240GB",
                storageType: "sata",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        },
        {
            title: "PC Intel Gamer Pro",
            desc: "Diseñada para jugar con excelente fluidez y realizar múltiples tareas a la vez.",
            specs: {
                cpu: "12400",
                motherboard: "H610M",
                ram: "DDR4 16GB 3200",
                ramQty: 1,
                storage: "M.2 NVME 512GB",
                storageType: "nvme",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        },
        {
            title: "PC Intel Ultimate Workstation",
            desc: "Rendimiento profesional absoluto para desarrollo, arquitectura y diseño multimedia.",
            specs: {
                cpu: "14700",
                motherboard: "B760M",
                ram: "DDR5 16GB",
                ramQty: 2,
                storage: "M.2 NVME 1TB",
                storageType: "nvme",
                case: "Magnum",
                cooling: "Stock",
                gpu: "Integrados",
                psu: "Fuente"
            }
        }
    ]
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
    } catch (error) {
        console.error("Error loading Excel:", error);
    }
}

function findBestMatch(items, query) {
    if (!items || items.length === 0) return null;
    const cleanQuery = query.toLowerCase().replace(/[^a-z0-9]/g, '');
    
    // First try: exact substring match
    for (const item of items) {
        const cleanName = item.name.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (cleanName.includes(cleanQuery)) {
            return item;
        }
    }
    
    // Second try: word matching score
    const parts = query.toLowerCase().split(' ');
    let bestItem = items[0];
    let maxMatches = 0;
    
    for (const item of items) {
        const cleanName = item.name.toLowerCase();
        let matches = 0;
        for (const part of parts) {
            if (part.length > 1 && cleanName.includes(part)) {
                matches++;
            }
        }
        if (matches > maxMatches) {
            maxMatches = matches;
            bestItem = item;
        }
    }
    return bestItem;
}

function renderPreconfiguredPcs(platform) {
    const container = document.getElementById('pc-list-container');
    if (!container) return;
    
    container.innerHTML = '';
    const pcs = predefinedPcs[platform];
    
    pcs.forEach((pc, idx) => {
        // Find best matches for prices
        const matchedCpu = findBestMatch(componentsData.cpu, pc.specs.cpu);
        const matchedMb = findBestMatch(componentsData.motherboard, pc.specs.motherboard);
        const matchedRam = findBestMatch(componentsData.ram, pc.specs.ram);
        
        let matchedStorage;
        if (pc.specs.storageType === 'nvme') {
            matchedStorage = findBestMatch(componentsData.storage.nvme, pc.specs.storage);
        } else {
            matchedStorage = findBestMatch(componentsData.storage.sata, pc.specs.storage);
        }
        
        const matchedCase = findBestMatch(componentsData.case, pc.specs.case);
        
        // Calculate Total
        let total = 0;
        if (matchedCpu) total += matchedCpu.price;
        if (matchedMb) total += matchedMb.price;
        if (matchedRam) total += matchedRam.price * pc.specs.ramQty;
        if (matchedStorage) total += matchedStorage.price;
        if (matchedCase) total += matchedCase.price;
        
        const card = document.createElement('div');
        card.className = 'card pc-card';
        
        // Render specs list
        card.innerHTML = `
            <div>
                <h3 class="pc-card-title">${pc.title}</h3>
                <p class="pc-card-desc">${pc.desc}</p>
                <ul class="pc-specs">
                    <li class="pc-spec-item">
                        <i class="ph-bold ph-cpu"></i>
                        <div>
                            <span class="pc-spec-label">Procesador:</span>
                            <span class="pc-spec-val">${matchedCpu ? matchedCpu.name : 'No disponible'}</span>
                        </div>
                    </li>
                    <li class="pc-spec-item">
                        <i class="ph-bold ph-square-logo"></i>
                        <div>
                            <span class="pc-spec-label">Motherboard:</span>
                            <span class="pc-spec-val">${matchedMb ? matchedMb.name : 'No disponible'}</span>
                        </div>
                    </li>
                    <li class="pc-spec-item">
                        <i class="ph-bold ph-square-half"></i>
                        <div>
                            <span class="pc-spec-label">Memoria RAM:</span>
                            <span class="pc-spec-val">${matchedRam ? `${matchedRam.name} (x${pc.specs.ramQty})` : 'No disponible'}</span>
                        </div>
                    </li>
                    <li class="pc-spec-item">
                        <i class="ph-bold ph-hard-drives"></i>
                        <div>
                            <span class="pc-spec-label">Almacenamiento:</span>
                            <span class="pc-spec-val">${matchedStorage ? matchedStorage.name : 'No disponible'}</span>
                        </div>
                    </li>
                    <li class="pc-spec-item">
                        <i class="ph-bold ph-frame-corners"></i>
                        <div>
                            <span class="pc-spec-label">Gabinete:</span>
                            <span class="pc-spec-val">${matchedCase ? matchedCase.name : 'No disponible'}</span>
                        </div>
                    </li>
                </ul>
            </div>
            <div>
                <div class="pc-price-box">
                    <span class="pc-price-label">Precio Estimado</span>
                    <span class="pc-price-val">$ ${total.toLocaleString('es-AR', {minimumFractionDigits: 2})}</span>
                </div>
                <button class="btn-primary" style="width:100%; text-align:center; justify-content:center;" onclick="customizePc(${idx}, '${platform}')">
                    <i class="ph-bold ph-gear-six"></i> Personalizar en Armador
                </button>
            </div>
        `;
        container.appendChild(card);
    });
}

function customizePc(idx, platform) {
    const pc = predefinedPcs[platform][idx];
    const matchedCpu = findBestMatch(componentsData.cpu, pc.specs.cpu);
    const matchedMb = findBestMatch(componentsData.motherboard, pc.specs.motherboard);
    const matchedRam = findBestMatch(componentsData.ram, pc.specs.ram);
    
    let matchedStorage;
    if (pc.specs.storageType === 'nvme') {
        matchedStorage = findBestMatch(componentsData.storage.nvme, pc.specs.storage);
    } else {
        matchedStorage = findBestMatch(componentsData.storage.sata, pc.specs.storage);
    }
    
    const matchedCase = findBestMatch(componentsData.case, pc.specs.case);
    
    // Construct Query String redirecting to index.html
    const params = new URLSearchParams();
    if (matchedCpu) {
        params.append('socket', matchedCpu.socket);
        params.append('cpu', matchedCpu.id);
    }
    if (matchedMb) params.append('motherboard', matchedMb.id);
    if (matchedRam) {
        params.append('ram', matchedRam.id);
        params.append('ramQty', pc.specs.ramQty);
    }
    if (matchedStorage) {
        params.append('storageType', pc.specs.storageType);
        params.append('storage', matchedStorage.id);
    }
    if (matchedCase) params.append('case', matchedCase.id);
    
    window.location.href = `index.html?${params.toString()}`;
}

async function init() {
    await loadExcelData();
    // Read platform from data-platform attribute in HTML
    const platform = document.documentElement.getAttribute('data-platform') || 'amd';
    renderPreconfiguredPcs(platform);
    
    // Theme toggle initialization (copy of index.html theme code)
    const html = document.documentElement;
    const btnLight = document.getElementById('btn-light-theme');
    const btnDark = document.getElementById('btn-dark-theme');
    
    if (btnLight && btnDark) {
        // Load preference from localStorage or default
        const savedTheme = localStorage.getItem('theme') || 'dark';
        html.setAttribute('data-theme', savedTheme);
        if (savedTheme === 'dark') {
            btnDark.classList.add('active');
            btnLight.classList.remove('active');
        } else {
            btnLight.classList.add('active');
            btnDark.classList.remove('active');
        }
        
        btnLight.addEventListener('click', () => {
            html.setAttribute('data-theme', 'light');
            localStorage.setItem('theme', 'light');
            btnLight.classList.add('active');
            btnDark.classList.remove('active');
        });
        btnDark.addEventListener('click', () => {
            html.setAttribute('data-theme', 'dark');
            localStorage.setItem('theme', 'dark');
            btnDark.classList.add('active');
            btnLight.classList.remove('active');
        });
    }
}

document.addEventListener('DOMContentLoaded', init);
