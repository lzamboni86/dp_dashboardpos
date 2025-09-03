// script.js - Versão Final Completa e Corrigida
// Recomendações importantes para inclusão no HTML (antes de script.js):
// <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
// <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

// =========================================================================
// Funções Auxiliares Globais (Acessíveis por toda a aplicação)
// =========================================================================

// --- Funções de Banco de Dados (IndexedDB) ---
/**
 * Inicializa e abre uma conexão com o banco de dados IndexedDB.
 * @param {string} dbName - Nome do banco de dados.
 * @param {string} storeName - Nome do object store a ser criado/usado.
 * @returns {Promise<IDBDatabase>} - Promessa que resolve com a instância do banco de dados.
 */
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

/**
 * Salva dados no object store do IndexedDB.
 * Os dados são armazenados sob a chave 'current', sobrescrevendo dados anteriores.
 * @param {IDBDatabase} db - Instância do banco de dados IndexedDB.
 * @param {string} storeName - Nome do object store.
 * @param {Array<Object>} data - Array de objetos a serem salvos.
 * @returns {Promise<void>} - Promessa que resolve quando os dados são salvos.
 */
async function saveData(db, storeName, data) {
    return new Promise((resolve, reject) => {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);

        // Otimização: limita o número de registros para evitar problemas de desempenho/armazenamento
        const optimizedData = data.slice(0, 5000);

        store.put(optimizedData, 'current'); // Armazena os dados sob a chave 'current'

        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
    });
}

/**
 * Carrega dados do object store do IndexedDB.
 * @param {IDBDatabase} db - Instância do banco de dados IndexedDB.
 * @param {string} storeName - Nome do object store.
 * @returns {Promise<Array<Object>>} - Promessa que resolve com os dados carregados ou um array vazio.
 */
async function loadData(db, storeName) {
    return new Promise((resolve) => {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const request = store.get('current');

        request.onsuccess = (e) => resolve(e.target.result || []); // Retorna dados ou array vazio
        request.onerror = (e) => {
            console.error("Erro ao carregar dados do IndexedDB:", e.target.error);
            resolve([]); // Em caso de erro, ainda resolve com um array vazio para evitar travamentos
        };
    });
}

// --- Função de Processamento de Excel (Requer a biblioteca SheetJS/XLSX) ---
/**
 * Processa um buffer de arquivo Excel, extraindo dados e formatando-os.
 * @param {ArrayBuffer} buffer - O buffer do arquivo Excel.
 * @returns {Array<Object>} - Array de objetos com os dados processados.
 * @throws {Error} - Se a planilha estiver vazia, sem cabeçalhos ou com colunas faltando.
 */
function processExcel(buffer) {
    try {
        const workbook = XLSX.read(buffer, {
            type: 'array',
            cellDates: true, // Tenta converter datas automaticamente
            sheetStubs: true // Inclui células vazias
        });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (data.length === 0) {
            throw new Error('Planilha vazia ou sem cabeçalhos.');
        }

        const headers = data[0].map(h => h ? String(h).trim() : '');
        const requiredColumns = [
            'Região', 'Etapa', 'Produto', 'PO Number',
            'Data da Criação', 'Status', 'Due Date',
            'Fornecedor', 'Rastreio', 'Last Update'
        ];

        const missingColumns = requiredColumns.filter(col => !headers.includes(col));
        if (missingColumns.length > 0) {
            throw new Error(`Colunas obrigatórias faltando: ${missingColumns.join(', ')}. Verifique a grafia exata.`);
        }

        // Mapeamento de nomes de colunas do Excel para chaves de objeto JS (boa prática)
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

        const processedRows = [];
        for (let i = 1; i < data.length; i++) { // Começa da segunda linha (após os cabeçalhos)
            const row = data[i];
            const rowData = {};
            requiredColumns.forEach(col => {
                const colIndex = headers.indexOf(col);
                let value = (colIndex !== -1 && row[colIndex] !== undefined) ? row[colIndex] : null;

                const mappedKey = columnMap[col] || col.toLowerCase().replace(/\s/g, '');

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

// --- Funções de Atualização da UI (Interface do Usuário) ---
/**
 * Atualiza todos os componentes da UI (cards, gráficos, tabela) com os novos dados.
 * @param {Object} uiElements - Objeto contendo referências aos elementos da UI.
 * @param {Array<Object>} data - Os dados processados do Excel.
 * @param {Object} chartInstances - Objeto para gerenciar instâncias de gráficos Chart.js.
 */
function updateUI(uiElements, data, chartInstances) {
    updateCards(data, uiElements.cardsContainer);
    updateCharts(data, uiElements.statusChartCanvas, uiElements.etapaChartCanvas, chartInstances);
    updateTable(data, uiElements.tableBody);
}

/**
 * Atualiza os cards de resumo com a contagem de itens por status.
 * @param {Array<Object>} data - Os dados processados.
 * @param {HTMLElement} cardsContainer - O container HTML para os cards.
 */
function updateCards(data, cardsContainer) {
    cardsContainer.innerHTML = ''; // Limpa os cards existentes

    const counts = data.reduce((acc, item) => {
        const statusKey = item.status?.toString().trim().toLowerCase() || 'desconhecido';
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

/**
 * Atualiza os gráficos de status e etapa.
 * @param {Array<Object>} data - Os dados processados.
 * @param {HTMLCanvasElement} statusChartCanvas - Canvas para o gráfico de status.
 * @param {HTMLCanvasElement} etapaChartCanvas - Canvas para o gráfico de etapa.
 * @param {Object} chartInstances - Objeto para gerenciar instâncias de gráficos Chart.js.
 */
function updateCharts(data, statusChartCanvas, etapaChartCanvas, chartInstances) {
    // Destrói instâncias de gráficos anteriores para evitar duplicação
    if (chartInstances.status) chartInstances.status.destroy();
    if (chartInstances.etapa) chartInstances.etapa.destroy();

    // Recria os gráficos com os novos dados
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

/**
 * Cria um gráfico de barras usando Chart.js.
 * @param {HTMLCanvasElement} canvas - O elemento canvas onde o gráfico será desenhado.
 * @param {Object} chartData - Objeto com os dados para o gráfico (ex: { 'Status A': 10, 'Status B': 20 }).
 * @param {string} label - O rótulo para o conjunto de dados do gráfico.
 * @returns {Chart} - A instância do gráfico Chart.js.
 */
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

/**
 * Atualiza a tabela de detalhes com os dados do Excel.
 * @param {Array<Object>} data - Os dados processados.
 * @param {HTMLElement} tableBody - O elemento tbody da tabela.
 */
function updateTable(data, tableBody) {
    tableBody.innerHTML = ""; // Limpa as linhas existentes da tabela

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
/**
 * Garante que um valor seja uma string e retorne uma string vazia se for null/undefined.
 * @param {*} value - O valor a ser formatado.
 * @returns {string} - O valor como string formatada.
 */
function safeText(value) {
    return value?.toString().trim() || "";
}

/**
 * Formata um valor de data para uma string de data localizada (pt-BR).
 * @param {*} dateValue - O valor da data (Date object, número ou string).
 * @returns {string} - A data formatada ou uma string vazia se inválida.
 */
function formatDate(dateValue) {
    return dateValue instanceof Date && !isNaN(dateValue) ? dateValue.toLocaleDateString("pt-BR") : "";
}

/**
 * Formata uma string de status (primeira letra maiúscula).
 * @param {string} status - A string de status.
 * @returns {string} - A string de status formatada.
 */
function formatStatus(status) {
    return status.charAt(0).toUpperCase() + status.slice(1);
}

/**
 * Converte um valor (número de Excel, string, ou Date object) para um Date object.
 * @param {*} value - O valor a ser convertido.
 * @returns {Date | null} - O Date object ou null se a conversão falhar.
 */
function parseDate(value) {
    if (value instanceof Date) {
        return value;
    }
    if (typeof value === 'number') {
        // Converte número de data do Excel para data JS (base 1900 vs 1970)
        return new Date(Math.round((value - 25569) * 86400 * 1000));
    }
    if (typeof value === 'string') {
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date; // Retorna null para strings de data inválidas
    }
    return null; // Retorna null para outros tipos
}

/**
 * Conta a ocorrência de valores em um campo específico dos dados.
 * @param {Array<Object>} data - O array de objetos.
 * @param {string} field - O nome do campo a ser contado.
 * @returns {Object} - Um objeto com as contagens (ex: { 'Valor A': 5, 'Valor B': 3 }).
 */
function getCountsByField(data, field) {
    return data.reduce((acc, item) => {
        const key = item?.[field]?.toString().trim() || "Não Informado";
        acc[key] = (acc[key] || 0) + 1;
        return acc;
    }, {});
}

// --- Tratamento de Erros ---
/**
 * Função genérica para lidar com erros, logando e exibindo um alerta.
 * @param {Error} error - O objeto de erro.
 */
function handleError(error) {
    console.error(`[${new Date().toISOString()}] ERRO:`, {
        message: error.message,
        stack: error.stack,
        type: error.name
    });
    alert(`Operação falhou: ${error.message.split(':').pop().trim()}`);
}

// --- Função para leitura de arquivo (mantida fora do DOMContentLoaded para globalidade) ---
/**
 * Lê um arquivo como um ArrayBuffer.
 * @param {File} file - O objeto File a ser lido.
 * @returns {Promise<ArrayBuffer>} - Promessa que resolve com o ArrayBuffer do arquivo.
 */
const readFile = (file) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (error) => reject(new Error('Falha na leitura do arquivo'));
        reader.readAsArrayBuffer(file);
    });
};

// =========================================================================
// Lógica Principal - Executada quando o DOM (Document Object Model) está completamente carregado
// =========================================================================
document.addEventListener('DOMContentLoaded', async () => {
    // Configuração do Banco de Dados
    const dbName = 'DashboardDB';
    const storeName = 'excelStore';
    let db = null;

    // 1. Inicialização do IndexedDB
    try {
        db = await initDB(dbName, storeName);
        console.log('IndexedDB inicializado com sucesso.');
    } catch (error) {
        console.error('Falha na inicialização do banco:', error);
        alert('Erro crítico: Armazenamento local indisponível. Verifique as configurações do navegador e tente novamente.');
        return; // Sai da função se o DB não puder ser inicializado, evitando erros posteriores
    }

    // 2. Referências aos Elementos da UI
    const ui = {
        excelUpload: document.getElementById('excelUpload'),
        cardsContainer: document.querySelector('.cards-container'),
        statusChartCanvas: document.getElementById('statusChart'),
        etapaChartCanvas: document.getElementById('etapaChart'),
        tableBody: document.querySelector('#dataTable tbody')
    };

    let chartInstances = { status: null, etapa: null };
    let excelData = []; // Variável para armazenar os dados do Excel na memória

    // 3. Função para lidar com o upload de arquivos Excel
    async function handleFileUpload(e) {
        const file = e.target.files[0];
        if (!file) return;

        try {
            const buffer = await readFile(file); // Lê o arquivo
            const processedData = processExcel(buffer); // Processa os dados do Excel
            await saveData(db, storeName, processedData); // Salva os dados no IndexedDB

            excelData = processedData; // Atualiza a variável local de dados
            updateUI(ui, excelData, chartInstances); // Atualiza a UI com os novos dados
            alert('Dados carregados e salvos com sucesso!');
            console.log('Novo Excel carregado e dados persistidos no IndexedDB.');

        } catch (error) {
            handleError(error); // Trata erros durante o upload/processamento
        }
    }

    // 4. Configurar Evento de Change para o Input de Arquivo
    ui.excelUpload.addEventListener('change', handleFileUpload);

    // 5. Carregar Dados Iniciais do IndexedDB e Atualizar a UI
    // Este bloco é crucial para a persistência ao reabrir a página
    try {
        excelData = await loadData(db, storeName); // Tenta carregar dados do IndexedDB
        if (excelData.length > 0) {
            updateUI(ui, excelData, chartInstances); // Atualiza a UI se houver dados
            console.log('Dados carregados do IndexedDB na inicialização:', excelData.length, 'registros.');
        } else {
            console.log('Nenhum dado encontrado no IndexedDB para carregar na inicialização.');
        }
    } catch (error) {
        console.error('Erro ao carregar dados do IndexedDB na inicialização:', error);
        alert('Erro ao carregar dados salvos. Tente carregar um novo arquivo Excel.');
    }
});