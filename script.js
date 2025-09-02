// script.js - Versão Final Completa e Corrigida

document.addEventListener('DOMContentLoaded', async () => {
    // Configuração do Banco de Dados
    const dbName = 'DashboardDB';
    const storeName = 'excelStore';
    let db = null;

    // Inicialização do IndexedDB
    try {
        db = await initDB(dbName, storeName);
    } catch (error) {
        console.error('Falha na inicialização do banco:', error);
        alert('Erro crítico: Armazenamento local indisponível');
        return;
    }

    // Elementos da UI
    const ui = {
        excelUpload: document.getElementById('excelUpload'),
        cardsContainer: document.querySelector('.cards-container'),
        statusChartCanvas: document.getElementById('statusChart'),
        etapaChartCanvas: document.getElementById('etapaChart'),
        tableBody: document.querySelector('#dataTable tbody')
    };

    // Função para leitura de arquivo (definida dentro do escopo DOMContentLoaded)
    const readFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (error) => reject(new Error('Falha na leitura do arquivo'));
            reader.readAsArrayBuffer(file);
        });
    };

    let chartInstances = { status: null, etapa: null };
    let excelData = []; // Irá conter os dados processados do Excel

    // Configurar eventos
    ui.excelUpload.addEventListener('change', handleFileUpload);

    // Carregar dados iniciais
    try {
        excelData = await loadData(db, storeName); // Chamada correta para loadData
        if (excelData.length > 0) updateUI(ui, excelData, chartInstances);
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
    }

    async function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        try {
            const buffer = await readFile(file);
            const processedData = processExcel(buffer); // processExcel definida globalmente
            await saveData(db, storeName, processedData); // saveData definida globalmente

            excelData = processedData; // Atualiza excelData global
            updateUI(ui, excelData, chartInstances); // Atualiza a UI com os novos dados
            alert('Dados carregados com sucesso!');
            
        } catch (error) {
            handleError(error); // handleError definida globalmente
        }
    }

    // Fim do DOMContentLoaded listener
}); // <<< ESSA CHAVE FECHA O document.addEventListener('DOMContentLoaded' ...


// =========================================================================
// Funções Auxiliares (Definidas FORA do DOMContentLoaded para acessibilidade global)
// =========================================================================

// --- Funções de Banco de Dados (IndexedDB) ---
async function initDB(dbName, storeName) {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(dbName, 1);

        request.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(storeName)) {
                db.createObjectStore(storeName);
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveData(db, storeName, data) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        
        // Exemplo de otimização de dados: limita a 5000 registros para evitar problemas de desempenho/armazenamento
        const optimizedData = data.slice(0, 5000); 

        store.put(optimizedData, 'current'); // Armazena os dados sob a chave 'current'

        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
    });
}

async function loadData(db, storeName) {
    return new Promise((resolve) => {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const request = store.get('current');

        request.onsuccess = (e) => resolve(e.target.result || []);
        request.onerror = () => resolve([]); // Retorna array vazio em caso de erro
    });
}


// --- Função de Processamento de Excel (Requer a biblioteca SheetJS/XLSX) ---
// Certifique-se de que <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
// esteja no seu HTML ANTES do seu script.js
function processExcel(buffer) {
    try {
        const workbook = XLSX.read(buffer, {
            type: 'array',
            cellDates: true, // Tenta converter datas automaticamente
            sheetStubs: true // Inclui células vazias
        });

        // Pega a primeira planilha
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Converte a planilha para um array de arrays, onde o primeiro array são os cabeçalhos
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (data.length === 0) {
            throw new Error('Planilha vazia ou sem cabeçalhos.');
        }

        const headers = data[0].map(h => h ? String(h).trim() : ''); // Garante que cabeçalhos sejam strings e trimados
        const requiredColumns = [
            'Região', 'Etapa', 'Produto', 'PO Number',
            'Data da Criação', 'Status', 'Due Date',
            'Fornecedor', 'Rastreio', 'Last Update'
        ];

        // Checagem básica para colunas obrigatórias
        const missingColumns = requiredColumns.filter(col => !headers.includes(col));
        if (missingColumns.length > 0) {
            throw new Error(`Colunas obrigatórias faltando: ${missingColumns.join(', ')}. Verifique a grafia exata.`);
        }

        const processedRows = [];
        for (let i = 1; i < data.length; i++) { // Começa da segunda linha (após os cabeçalhos)
            const row = data[i];
            const rowData = {};
            requiredColumns.forEach(col => {
                const colIndex = headers.indexOf(col);
                let value = (colIndex !== -1 && row[colIndex] !== undefined) ? row[colIndex] : null;

                // Mapeia o nome da coluna do Excel para a chave do objeto JS
                const mappedKey = columnMap[col] || col.toLowerCase().replace(/\s/g, '');
                
                // Tratamento especial para datas que podem vir como números ou strings
                if (col.includes('Date')) {
                    rowData[mappedKey] = parseDate(value);
                } else {
                    rowData[mappedKey] = value;
                }
            });
            processedRows.push(rowData);
        }
        return processedRows;
    } catch (error) {
        console.error('Erro no processamento do Excel:', error);
        throw new Error(`Erro ao processar o arquivo Excel: ${error.message}`);
    }
}

// Mapeamento de nomes de colunas do Excel para chaves de objeto JS (opcional, mas boa prática)
const columnMap = {
    'Região': 'regiao',
    'Etapa': 'etapa',
    'Produto': 'produto',
    'PO Number': 'poNumber',
    'Data da Criação': 'criacao',
    'Status': 'status',
    'Due Date': 'dueDate',
    'Fornecedor': 'fornecedor',
    'Rastreio': 'rastreio',
    'Last Update': 'lastUpdate'
};


// --- Funções de Atualização da UI ---
// Certifique-se de que <script src="https://cdn.jsdelivr.net/npm/chart.js@3.7.0/dist/chart.min.js"></script>
// esteja no seu HTML ANTES do seu script.js
function updateUI(uiElements, data, chartInstances) {
    updateCards(data, uiElements.cardsContainer);
    updateCharts(data, uiElements.statusChartCanvas, uiElements.etapaChartCanvas, chartInstances);
    updateTable(data, uiElements.tableBody);
}

function updateCards(data, cardsContainer) {
    cardsContainer.innerHTML = '';
    
    const counts = data.reduce((acc, item) => {
        const statusKey = item.status?.toString().trim().toLowerCase() || 'desconhecido'; // Usa toString para segurança
        acc[statusKey] = (acc[statusKey] || 0) + 1;
        return acc;
    }, {});

    Object.keys(counts).forEach(statusKey => {
        const cardHtml = `
            <div class="card ${statusKey}">
                <h3>${formatStatus(statusKey)}</h3>
                <span>${counts[statusKey]}</span>
            </div>
        `;
        cardsContainer.insertAdjacentHTML('beforeend', cardHtml);
    });
}

function updateCharts(data, statusChartCanvas, etapaChartCanvas, chartInstances) {
    // Destroi instâncias de gráficos anteriores se existirem
    if (chartInstances.status) chartInstances.status.destroy();
    if (chartInstances.etapa) chartInstances.etapa.destroy();

    // Recria os gráficos
    chartInstances.status = createBarChart(
        statusChartCanvas,
        getCountsByField(data, "status"),
        "Distribuição por Status"
    );

    chartInstances.etapa = createBarChart(
        etapaChartCanvas,
        getCountsByField(data, "etapa"),
        "Distribuição por Etapa"
    );
}

function createBarChart(canvas, chartData, label) {
    const ctx = canvas.getContext("2d");
    return new Chart(ctx, {
        type: "bar",
        data: {
            labels: Object.keys(chartData),
            datasets: [{
                label: label,
                data: Object.values(chartData),
                backgroundColor: "#2196F3", // Cor padrão para barras
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true }
            }
        }
    });
}

function updateTable(data, tableBody) {
    tableBody.innerHTML = "";
    
    data.forEach(item => {
        const rowHtml = `
            <tr>
                <td>${safeText(item.regiao)}</td>
                <td>${safeText(item.etapa)}</td>
                <td>${safeText(item.produto)}</td>
                <td>${safeText(item.poNumber)}</td>
                <td>${formatDate(item.criacao)}</td>
                <td>${safeText(item.status)}</td>
                <td>${formatDate(item.dueDate)}</td>
                <td>${safeText(item.fornecedor)}</td>
                <td>${safeText(item.rastreio)}</td>
                <td>${formatDate(item.lastUpdate)}</td>
            </tr>
        `;
        tableBody.insertAdjacentHTML('beforeend', rowHtml);
    });
}


// --- Funções Utilitárias ---
function safeText(value) {
    return value?.toString().trim() || ""; // Garante que o valor é string e retorna vazio se null/undefined
}

function formatDate(dateValue) {
    // Garante que dateValue é um objeto Date válido antes de formatar
    return dateValue instanceof Date && !isNaN(dateValue) ? dateValue.toLocaleDateString("pt-BR") : "";
}

function formatStatus(status) {
    return status.charAt(0).toUpperCase() + status.slice(1);
}

function parseDate(value) {
    if (value instanceof Date) {
        return value;
    }
    if (typeof value === 'number') {
        // Converte número de data do Excel para data JS
        // O número 25569 é a diferença de dias entre 1900-01-01 (base do Excel) e 1970-01-01 (base do JS)
        return new Date(Math.round((value - 25569) * 86400 * 1000));
    }
    if (typeof value === 'string') {
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date; // Retorna null para strings de data inválidas
    }
    return null; // Retorna null para outros tipos
}

function getCountsByField(data, field) {
    return data.reduce((acc, item) => {
        const key = item?.[field]?.toString().trim() || "Não Informado";
        acc[key] = (acc[key] || 0) + 1;
        return acc;
    }, {});
}

// --- Tratamento de Erros ---
function handleError(error) {
    console.error(`[${new Date().toISOString()}] ERRO:`, {
        message: error.message,
        stack: error.stack,
        type: error.name
    });

    // Exibe apenas a parte mais relevante da mensagem de erro
    alert(`Operação falhou: ${error.message.split(':').pop().trim()}`);
}

// --- Logout (Exemplo, se aplicável ao seu projeto) ---
// window.logout = () => {
//    localStorage.clear(); // Ou limpe itens específicos
//    window.location.href = "index.html";
// };