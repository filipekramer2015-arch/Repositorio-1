// Variáveis Globais de Estado
let rawData = [];
let filteredData = [];
let dataTableInstance = null;
let charts = {};
let columnAnalysis = { categorias: [], numericas: [] };

// Toggle Sidebar Mobile/Desktop
document.getElementById('menu-toggle').addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('toggled');
});

// Toggle Modo Claro/Escuro
document.getElementById('themeToggle').addEventListener('click', function() {
    const body = document.body;
    body.classList.toggle('dark-mode');
    body.classList.toggle('light-mode');
    
    const isDark = body.classList.contains('dark-mode');
    this.innerHTML = isDark ? '<i class="fas fa-sun me-2"></i>Modo Claro' : '<i class="fas fa-moon me-2"></i>Modo Escuro';
    
    Chart.defaults.color = isDark ? '#b2bec3' : '#636e72';
    Chart.defaults.borderColor = isDark ? '#2d3436' : '#e9ecef';
    updateAllCharts(); // Re-renderiza gráficos para ajustar cores das fontes
});

// Exportar PDF
document.getElementById('exportPdf').addEventListener('click', () => {
    const element = document.getElementById('dashboard-container');
    const opt = {
        margin:       10,
        filename:     'Dashboard_Report.pdf',
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'landscape' }
    };
    html2pdf().set(opt).from(element).save();
});

// Leitura Automática do Arquivo Excel
document.getElementById('excelInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Pega a primeira aba da planilha
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Converte para JSON
        rawData = XLSX.utils.sheet_to_json(worksheet);
        filteredData = [...rawData];

        if (rawData.length > 0) {
            document.getElementById('uploadPrompt').style.display = 'none';
            document.getElementById('mainDashboard').style.display = 'block';
            analyzeDataStructure();
            populateFilters();
            renderDashboard();
        } else {
            alert("A planilha está vazia.");
        }
    };
    reader.readAsArrayBuffer(file);
});

// INTELIGÊNCIA: Analisa os dados para descobrir quais colunas são números e quais são texto (categorias)
function analyzeDataStructure() {
    if (rawData.length === 0) return;
    
    columnAnalysis.categorias = [];
    columnAnalysis.numericas = [];
    
    const sample = rawData[0];
    for (const key in sample) {
        // Verifica se a maioria dos dados daquela coluna é número
        let isNumeric = rawData.slice(0, 10).every(row => {
            return !isNaN(parseFloat(row[key])) && isFinite(row[key]);
        });

        // Ignora IDs ou Códigos como colunas matemáticas
        if (isNumeric && !key.toLowerCase().includes('id') && !key.toLowerCase().includes('código')) {
            columnAnalysis.numericas.push(key);
        } else {
            columnAnalysis.categorias.push(key);
        }
    }
}

// Renderiza tudo (KPIs, Gráficos, Tabela) com base nos dados filtrados
function renderDashboard() {
    updateKPIs();
    updateAllCharts();
    renderTable();
}

// Cálculos dos KPIs
function updateKPIs() {
    if (filteredData.length === 0) return;

    // KPI 1: Total Registros
    document.getElementById('kpi-records').innerText = filteredData.length;

    // Se houver colunas numéricas, pega a primeira como principal para cálculos
    if (columnAnalysis.numericas.length > 0) {
        const mainMetric = columnAnalysis.numericas[0];
        const qtyMetric = columnAnalysis.numericas[1] || mainMetric; // Segunda coluna numérica para "Quantidade"

        let total = 0;
        let totalQty = 0;
        let min = Infinity;
        let max = -Infinity;

        filteredData.forEach(row => {
            let val = parseFloat(row[mainMetric]) || 0;
            let qty = parseFloat(row[qtyMetric]) || 0;
            
            total += val;
            totalQty += qty;
            if (val > max) max = val;
            if (val < min) min = val;
        });

        let avg = total / filteredData.length;

        // Formatação Monetária/Numérica
        const formatBRL = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' });
        const formatNum = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 2 });

        document.getElementById('kpi-total').innerText = formatBRL.format(total);
        document.getElementById('kpi-qty').innerText = formatNum.format(totalQty);
        document.getElementById('kpi-avg').innerText = formatNum.format(avg);
        document.getElementById('kpi-max').innerText = formatBRL.format(max);
        document.getElementById('kpi-min').innerText = formatBRL.format(min);
    }
}

// Lógica de Geração Automática dos Gráficos
function updateAllCharts() {
    if (filteredData.length === 0) return;

    // Seleciona de forma inteligente os eixos
    const dim1 = columnAnalysis.categorias[0] || 'Categoria';
    const dim2 = columnAnalysis.categorias[1] || dim1;
    const metric1 = columnAnalysis.numericas[0] || 'Valor';
    
    // Agrupa Dados
    const groupedByDim1 = aggregateData(filteredData, dim1, metric1);
    const groupedByDim2 = aggregateData(filteredData, dim2, metric1);

    // Paleta de Cores Corporativas Modernas
    const colors = ['#0d6efd', '#20c997', '#ffc107', '#dc3545', '#6f42c1', '#fd7e14'];

    createChart('barChart', 'bar', Object.keys(groupedByDim1), Object.values(groupedByDim1), 'Total por ' + dim1, colors);
    createChart('pieChart', 'pie', Object.keys(groupedByDim2), Object.values(groupedByDim2), 'Proporção por ' + dim2, colors);
    createChart('doughnutChart', 'doughnut', Object.keys(groupedByDim1), Object.values(groupedByDim1), 'Distribuição', colors);
    
    // Gráfico de Linha/Área (Utiliza a primeira coluna categórica assumindo ordem cronológica/fase)
    createChart('lineChart', 'line', Object.keys(groupedByDim1), Object.values(groupedByDim1), 'Evolução', '#0d6efd', true);
}

// Função Auxiliar para Criar/Atualizar Gráficos via Chart.js
function createChart(canvasId, type, labels, data, title, colorInfo, isArea = false) {
    const ctx = document.getElementById(canvasId).getContext('2d');
    
    if (charts[canvasId]) {
        charts[canvasId].destroy(); // Previne sobreposição e duplicação
    }

    const config = {
        type: type,
        data: {
            labels: labels,
            datasets: [{
                label: title,
                data: data,
                backgroundColor: isArea ? 'rgba(13, 110, 253, 0.2)' : colorInfo,
                borderColor: isArea ? '#0d6efd' : (type === 'line' ? '#0d6efd' : '#fff'),
                borderWidth: type === 'pie' || type === 'doughnut' ? 2 : 1,
                fill: isArea,
                tension: 0.4 // Curva suave para linha
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: type === 'pie' || type === 'doughnut' ? 'right' : 'top' }
            }
        }
    };

    charts[canvasId] = new Chart(ctx, config);
}

// Função Auxiliar de Agrupamento Matemático de Dados (Soma por Categoria)
function aggregateData(data, groupByKey, sumKey) {
    return data.reduce((acc, curr) => {
        let group = curr[groupByKey] || 'Outros';
        let val = parseFloat(curr[sumKey]) || 0;
        acc[group] = (acc[group] || 0) + val;
        return acc;
    }, {});
}

// Tabela Profissional usando DataTables
function renderTable() {
    if (dataTableInstance) {
        dataTableInstance.destroy();
        document.getElementById('tableHead').innerHTML = '';
        document.getElementById('tableBody').innerHTML = '';
    }

    if (filteredData.length === 0) return;

    const keys = Object.keys(filteredData[0]);
    
    // Cabeçalho
    const thead = document.getElementById('tableHead');
    keys.forEach(key => {
        let th = document.createElement('th');
        th.innerText = key;
        thead.appendChild(th);
    });

    // Corpo
    const tbody = document.getElementById('tableBody');
    filteredData.forEach(row => {
        let tr = document.createElement('tr');
        keys.forEach(key => {
            let td = document.createElement('td');
            // Tenta formatar se for número grande
            if(!isNaN(row[key]) && row[key] !== '') {
                td.innerText = Number.isInteger(row[key]) ? row[key] : parseFloat(row[key]).toFixed(2);
            } else {
                td.innerText = row[key] || '-';
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    // Inicializa DataTable (Ordenação, Paginação nativa)
    dataTableInstance = $('#dataTable').DataTable({
        language: { url: '//cdn.datatables.net/plug-ins/1.13.4/i18n/pt-BR.json' },
        pageLength: 5,
        lengthMenu: [5, 10, 25, 50]
    });
}

// Sistema de Filtros Inteligentes
function populateFilters() {
    if (columnAnalysis.categorias.length >= 2) {
        // Povoa filtro de Categoria (Dimensão 1)
        const catSelect = document.getElementById('filterCategory');
        const catSet = new Set(rawData.map(item => item[columnAnalysis.categorias[0]]));
        catSelect.innerHTML = '<option value="">Todas</option>';
        catSet.forEach(val => { if(val) catSelect.innerHTML += `<option value="${val}">${val}</option>`; });

        // Povoa filtro de Período/Referência (Dimensão 2 ou Data se existir)
        const dateSelect = document.getElementById('filterDate');
        const dateSet = new Set(rawData.map(item => item[columnAnalysis.categorias[1]]));
        dateSelect.innerHTML = '<option value="">Todos</option>';
        dateSet.forEach(val => { if(val) dateSelect.innerHTML += `<option value="${val}">${val}</option>`; });
    }
}

// Execução dos Filtros
function applyFilters() {
    const term = document.getElementById('globalSearch').value.toLowerCase();
    const catVal = document.getElementById('filterCategory').value;
    const dateVal = document.getElementById('filterDate').value;
    
    const dim1 = columnAnalysis.categorias[0];
    const dim2 = columnAnalysis.categorias[1];

    filteredData = rawData.filter(row => {
        let matchTerm = Object.values(row).some(val => String(val).toLowerCase().includes(term));
        let matchCat = catVal === "" || row[dim1] == catVal;
        let matchDate = dateVal === "" || row[dim2] == dateVal;
        return matchTerm && matchCat && matchDate;
    });

    renderDashboard();
}

// Listeners dos Filtros
document.getElementById('globalSearch').addEventListener('input', applyFilters);
document.getElementById('filterCategory').addEventListener('change', applyFilters);
document.getElementById('filterDate').addEventListener('change', applyFilters);

document.getElementById('clearFilters').addEventListener('click', () => {
    document.getElementById('globalSearch').value = '';
    document.getElementById('filterCategory').value = '';
    document.getElementById('filterDate').value = '';
    applyFilters();
});