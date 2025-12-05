// =================================================================
// CONFIGURAÇÃO GLOBAL
// =================================================================
const APPS_SCRIPT_API_URL = "https://script.google.com/macros/s/AKfycbxKRogXSYUx7Fmay-OO98v1sc1Z8cqf_nuwPy9fKN1aZJGWHkjZ6BCpcxLYqutnK9Qweg/exec";

// URL Direta (Servidor de Conteúdo Google) - Mais estável para Scripts
const FIXED_LOGO_ID = "1ipB03pEAEsV_u0jBzLF27qEjgAoAX4U8";
// Usamos o domínio lh3.googleusercontent.com pois ele aceita melhor conexões externas (CORS)
const FIXED_LOGO_URL = `https://lh3.googleusercontent.com/d/${FIXED_LOGO_ID}`;

let AppState = {
    selectedSpreadsheetId: null,
    selectedCondoName: null,
    logoData: null, 
    medicoes: [],
    previewData: {} 
};

// =================================================================
// CLASSES DE UTILIDADE
// =================================================================
class NotificationManager {
    static show(message, type = 'info', duration = 5000) {
        const container = document.getElementById('toast-container');
        if (!container) return;
        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        
        let icon = 'info-circle';
        if (type === 'success') icon = 'check-circle';
        if (type === 'error') icon = 'exclamation-circle';
        if (type === 'warning') icon = 'exclamation-triangle';
        
        toast.innerHTML = `
            <div class="flex items-center gap-3">
                <i class="fas fa-${icon} text-lg opacity-90"></i>
                <span>${message}</span>
            </div>
            <button onclick="this.parentElement.remove()" class="text-white hover:text-gray-200 transition-colors">
                <i class="fas fa-times"></i>
            </button>
        `;
        container.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 100);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }
    static success(message) { this.show(message, 'success'); }
    static error(message) { this.show(message, 'error', 8000); }
    static warning(message) { this.show(message, 'warning'); }
}

class ProgressManager {
    static show() { 
        const el = document.getElementById('progress-container');
        if(el) {
            el.classList.remove('hidden'); 
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    }
    static hide() { document.getElementById('progress-container')?.classList.add('hidden'); }
    static update(percent, text = '') {
        const fill = document.getElementById('progress-fill');
        const textEl = document.getElementById('progress-text');
        const labelEl = document.getElementById('progress-label');
        if (fill) fill.style.width = `${percent}%`;
        if (textEl) textEl.textContent = `${Math.round(percent)}%`;
        if (labelEl) labelEl.innerHTML = `<i class="fas fa-cog fa-spin mr-2"></i> ${text}`;
    }
}

// =================================================================
// CLASSE PRINCIPAL DA APLICAÇÃO
// =================================================================
class DashboardApp {
    constructor() {
        this.initializeElements();
        this.attachEventListeners();
        this.loadCondominios();
        this.loadFixedLogo(); 
        
        const yearEl = document.getElementById('year');
        const yearLoginEl = document.getElementById('year-login');
        const currentYear = new Date().getFullYear();
        if(yearEl) yearEl.textContent = currentYear;
        if(yearLoginEl) yearLoginEl.textContent = currentYear;
    }

    initializeElements() {
        this.elements = {
            condoSelectorScreen: document.getElementById('condo-selector-screen'),
            mainDashboardScreen: document.getElementById('main-dashboard-screen'),
            condoList: document.getElementById('condo-list'),
            dashboardTitle: document.getElementById('dashboard-title'),
            backToSelectorBtn: document.getElementById('back-to-selector'),
            fileInput: document.getElementById('fileInput'),
            fileFeedback: document.getElementById('file-feedback'),
            periodoDe: document.getElementById('periodo-de'),
            periodoAte: document.getElementById('periodo-ate'),
            proximaLeitura: document.getElementById('proxima-leitura'),
            tarifaEnergia: document.getElementById('tarifa-energia'),
            taxaGestao: document.getElementById('taxa-gestao'),
            rateioAreaComum: document.getElementById('rateio-area-comum'),
            processButton: document.getElementById('process-button'),
            previewButton: document.getElementById('preview-button'),
            downloadLinksSection: document.getElementById('download-links-section'),
            downloadLinksContainer: document.getElementById('download-links-container'),
            reportContainer: document.getElementById('reportContainer'),
            previewContainer: document.getElementById('previewContainer'),
            closePreviewButton: document.getElementById('closePreviewButton'),
            previewTabs: document.getElementById('preview-tabs'),
            previewTabContent: document.getElementById('preview-tab-content'),
            logoStatus: document.getElementById('logo-loading-status')
        };
    }

    attachEventListeners() {
        this.elements.backToSelectorBtn.addEventListener('click', () => this.showSelectorScreen());
        this.elements.processButton.addEventListener('click', () => this.handleProcessRequest());
        this.elements.previewButton.addEventListener('click', () => this.showPreview());
        this.elements.closePreviewButton.addEventListener('click', () => this.elements.previewContainer.classList.add('hidden'));
        this.elements.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        this.elements.reportContainer.addEventListener('input', (e) => this.handleCellEdit(e));
        this.elements.reportContainer.addEventListener('change', (e) => this.handleCheckboxChange(e));

        const formInputs = [this.elements.periodoDe, this.elements.periodoAte, this.elements.tarifaEnergia];
        formInputs.forEach(input => input.addEventListener('input', () => this.validateForm()));
    }

    // --- CORREÇÃO DO LOGO ---
    // Usa a técnica de criar uma Imagem HTML e desenhar num Canvas
    // Isso contorna problemas de CORS que o fetch() simples enfrenta
    loadFixedLogo() {
        const img = new Image();
        img.crossOrigin = "Anonymous"; // Permite carregar de outro domínio
        img.src = FIXED_LOGO_URL;

        img.onload = () => {
            try {
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                
                // Extrai os dados em Base64
                const dataURL = canvas.toDataURL('image/png');
                // Remove o prefixo para enviar apenas os bytes
                AppState.logoData = dataURL.replace(/^data:image\/(png|jpg|jpeg);base64,/, "");

                if(this.elements.logoStatus) {
                    this.elements.logoStatus.innerHTML = `<i class="fas fa-check-circle text-green-500 mr-1"></i> Carregado com sucesso`;
                    this.elements.logoStatus.className = "mt-2 text-xs text-green-600 font-bold flex items-center";
                }
            } catch (e) {
                console.error("Erro ao processar logo:", e);
                this.handleLogoError();
            }
        };

        img.onerror = () => {
            console.error("Erro ao baixar a imagem do logo.");
            this.handleLogoError();
        };
    }

    handleLogoError() {
        if(this.elements.logoStatus) {
            this.elements.logoStatus.innerHTML = `<i class="fas fa-exclamation-triangle text-amber-500 mr-1"></i> Usando padrão do sistema`;
            this.elements.logoStatus.className = "mt-2 text-xs text-amber-600 font-bold flex items-center";
        }
        AppState.logoData = null; // Backend usará fallback se necessário
    }

    async loadCondominios() {
        if (!APPS_SCRIPT_API_URL || APPS_SCRIPT_API_URL.includes("COLE_AQUI")) {
            this.elements.condoList.innerHTML = `<div class="col-span-full text-center p-10 bg-white rounded-lg shadow-sm border border-red-100"><i class="fas fa-exclamation-triangle text-red-500 text-4xl mb-4"></i><h3 class="text-xl font-bold text-red-700">API não configurada</h3><p class="text-slate-600 mt-2">A URL da API do Google Apps Script não foi definida no arquivo <strong>script.js</strong>.</p></div>`;
            return;
        }
        try {
            const request = new Request(APPS_SCRIPT_API_URL, { method: 'GET', mode: 'cors', redirect: 'follow' });
            const response = await fetch(request);
            if (!response.ok) throw new Error(`Erro de rede: ${response.status} ${response.statusText}`);
            const result = await response.json();
            if (!result.success) throw new Error(result.message);
            this.displayCondoCards(result.data);
        } catch (error) {
            this.elements.condoList.innerHTML = `<div class="col-span-full text-center p-10 bg-white rounded-lg shadow-sm border border-red-100"><i class="fas fa-wifi text-red-400 text-4xl mb-4"></i><h3 class="text-lg font-bold text-slate-700">Falha na Conexão</h3><p class="text-slate-500 mt-2">Não foi possível buscar a lista.</p><p class="text-xs text-slate-400 mt-4 font-mono bg-slate-100 p-2 rounded inline-block">${error.message}</p></div>`;
        }
    }

    displayCondoCards(condominios) {
        this.elements.condoList.innerHTML = '';
        if (!condominios || condominios.length === 0) {
            this.elements.condoList.innerHTML = `<div class="col-span-full text-center p-10 bg-white rounded-lg shadow"><i class="fas fa-folder-open text-yellow-400 text-4xl mb-4"></i><h3 class="text-xl font-bold text-slate-700">Nenhum condomínio encontrado</h3><p class="text-slate-500 mt-2">Verifique a pasta no Google Drive.</p></div>`;
            return;
        }
        condominios.forEach((condo, index) => {
            const card = document.createElement('div');
            card.style.animationDelay = `${index * 100}ms`;
            card.className = 'bg-white rounded-lg shadow-sm p-6 cursor-pointer hover:shadow-lg hover:-translate-y-1 transition-all duration-300 border border-slate-200 group slide-up relative overflow-hidden';
            card.innerHTML = `
                <div class="absolute top-0 left-0 w-1 h-full bg-slate-200 group-hover:bg-tech-blue transition-colors"></div>
                <div class="flex items-center justify-between mb-4">
                    <div class="w-12 h-12 bg-slate-50 rounded-lg flex items-center justify-center text-slate-400 group-hover:bg-tech-blue group-hover:text-white transition-colors">
                        <i class="fas fa-building text-xl"></i>
                    </div>
                    <i class="fas fa-chevron-right text-slate-300 group-hover:text-tech-blue transition-colors"></i>
                </div>
                <h3 class="text-lg font-bold text-slate-800 group-hover:text-tech-blue transition-colors">${condo.name}</h3>
                <p class="text-xs text-slate-400 mt-2">Ref: ${condo.id.substring(0, 8)}...</p>
            `;
            card.addEventListener('click', () => this.selectCondo(condo.id, condo.name));
            this.elements.condoList.appendChild(card);
        });
    }

    selectCondo(spreadsheetId, condoName) {
        AppState.selectedSpreadsheetId = spreadsheetId;
        AppState.selectedCondoName = condoName;
        this.elements.dashboardTitle.innerHTML = `${condoName}`;
        this.showDashboardScreen();
    }

    showDashboardScreen() {
        this.elements.condoSelectorScreen.classList.add('hidden');
        this.elements.mainDashboardScreen.classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
        this.resetDashboard();
    }

    showSelectorScreen() {
        this.elements.mainDashboardScreen.classList.add('hidden');
        this.elements.condoSelectorScreen.classList.remove('hidden');
        AppState.selectedSpreadsheetId = null;
    }

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (!file) { this.elements.fileFeedback.innerHTML = ''; return; }
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            NotificationManager.error("Formato de arquivo inválido. Use .xlsx");
            this.elements.fileFeedback.innerHTML = `<div class="text-red-500 font-medium bg-red-50 p-2 rounded text-xs"><i class="fas fa-times-circle mr-1"></i> Inválido</div>`;
            this.elements.fileInput.value = '';
            return;
        }
        this.elements.fileFeedback.innerHTML = `<div class="text-green-700 font-medium bg-green-50 p-2 rounded text-xs flex items-center justify-center"><i class="fas fa-check-circle mr-2"></i> ${file.name}</div>`;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                this.processAndDisplayTable(jsonData);
            } catch (error) {
                NotificationManager.error("Erro ao ler o arquivo .xlsx.");
                console.error(error);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    processAndDisplayTable(jsonData) {
        if (jsonData.length === 0) {
            NotificationManager.warning("A planilha está vazia.");
            AppState.medicoes = [];
            this.renderTable();
            return;
        }
        const normalizeHeader = (h) => String(h).trim().toLowerCase().replace(/\s+/g, '_');
        const normalizedData = jsonData.map(row => {
            const newRow = {};
            for (const key in row) {
                newRow[normalizeHeader(key)] = row[key];
            }
            return newRow;
        });
        AppState.medicoes = normalizedData.map(row => ({
            unidade: row.unidade || 'N/A',
            leitura_anterior: parseFloat(row.leitura_anterior) || 0,
            leitura_atual: parseFloat(row.leitura_atual) || 0,
            isCommonArea: ['true', 'sim', '1', 's', 'verdadeiro'].includes(String(row.area_comum).trim().toLowerCase())
        }));
        this.renderTable();
    }

    renderTable() {
        if (AppState.medicoes.length === 0) {
            this.elements.reportContainer.innerHTML = `<div class="flex flex-col items-center justify-center py-20 text-slate-400 bg-slate-50/30"><div class="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mb-4 text-slate-300"><i class="fas fa-table text-2xl"></i></div><p>Aguardando arquivo...</p></div>`;
            this.validateForm();
            return;
        }
        let tableHTML = `<div class="table-container bg-white"><table class="min-w-full divide-y divide-gray-200">
            <thead>
                <tr>
                    <th class="px-6 py-4 text-left text-xs font-bold text-tech-blue uppercase tracking-wider bg-slate-50">Unidade</th>
                    <th class="px-6 py-4 text-left text-xs font-bold text-tech-blue uppercase tracking-wider bg-slate-50">Leitura Anterior</th>
                    <th class="px-6 py-4 text-left text-xs font-bold text-tech-blue uppercase tracking-wider bg-slate-50">Leitura Atual</th>
                    <th class="px-6 py-4 text-left text-xs font-bold text-tech-blue uppercase tracking-wider bg-slate-50">Consumo (m³)</th>
                    <th class="px-6 py-4 text-center text-xs font-bold text-tech-blue uppercase tracking-wider bg-slate-50">Área Comum</th>
                </tr>
            </thead>
            <tbody class="bg-white divide-y divide-gray-100">`;
        
        AppState.medicoes.forEach((row, index) => {
            const consumo = (row.leitura_atual - row.leitura_anterior).toFixed(3);
            const isNegative = consumo < 0;
            
            tableHTML += `<tr data-index="${index}" class="hover:bg-tech-light transition-colors ${isNegative ? 'bg-red-50' : ''}">
                <td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-slate-700">
                    <span class="bg-slate-100 text-slate-600 px-2 py-1 rounded text-xs font-bold border border-slate-200">${row.unidade}</span>
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-slate-600 editable-cell relative group" contenteditable="true" data-field="leitura_anterior" title="Clique para editar">
                    ${row.leitura_anterior}
                    <i class="fas fa-pen text-[10px] text-tech-blue absolute right-2 top-4 opacity-0 group-hover:opacity-100 transition-opacity"></i>
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-sm text-slate-600 editable-cell relative group" contenteditable="true" data-field="leitura_atual" title="Clique para editar">
                    ${row.leitura_atual}
                    <i class="fas fa-pen text-[10px] text-tech-blue absolute right-2 top-4 opacity-0 group-hover:opacity-100 transition-opacity"></i>
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-sm font-bold consumo-cell ${isNegative ? 'text-red-600' : 'text-green-600'}">
                    ${consumo}
                </td>
                <td class="px-6 py-4 whitespace-nowrap text-center">
                    <input type="checkbox" class="common-area-checkbox" ${row.isCommonArea ? 'checked' : ''}>
                </td>
            </tr>`;
        });
        tableHTML += `</tbody></table></div>`;
        this.elements.reportContainer.innerHTML = tableHTML;
        this.validateForm();
    }

    handleCellEdit(event) {
        const cell = event.target;
        if (!cell.classList.contains('editable-cell')) return;
        const rowEl = cell.closest('tr');
        const index = parseInt(rowEl.dataset.index);
        const field = cell.dataset.field;
        const newValue = parseFloat(cell.textContent) || 0;
        AppState.medicoes[index][field] = newValue;
        const updatedRow = AppState.medicoes[index];
        const consumo = updatedRow.leitura_atual - updatedRow.leitura_anterior;
        const consumoCell = rowEl.querySelector('.consumo-cell');
        consumoCell.textContent = consumo.toFixed(3);
        
        if (consumo < 0) {
            rowEl.classList.add('bg-red-50');
            consumoCell.classList.remove('text-green-600');
            consumoCell.classList.add('text-red-600');
        } else {
            rowEl.classList.remove('bg-red-50');
            consumoCell.classList.remove('text-red-600');
            consumoCell.classList.add('text-green-600');
        }
    }

    handleCheckboxChange(event) {
        const checkbox = event.target;
        if (checkbox.type !== 'checkbox' || !checkbox.classList.contains('common-area-checkbox')) return;
        const rowEl = checkbox.closest('tr');
        if (!rowEl) return;
        const index = parseInt(rowEl.dataset.index);
        if (isNaN(index)) return;
        AppState.medicoes[index].isCommonArea = checkbox.checked;
    }

    validateForm() {
        const fileOk = AppState.medicoes.length > 0;
        const datesOk = this.elements.periodoDe.value && this.elements.periodoAte.value;
        const tarifaOk = this.elements.tarifaEnergia.value;
        const isValid = fileOk && datesOk && tarifaOk;
        this.elements.processButton.disabled = !isValid;
        this.elements.previewButton.disabled = !isValid;
        return isValid;
    }

    resetDashboard() {
        AppState.medicoes = [];
        AppState.previewData = {};
        this.elements.fileInput.value = '';
        this.elements.fileFeedback.innerHTML = '';
        this.elements.downloadLinksSection.classList.add('hidden');
        this.elements.downloadLinksContainer.innerHTML = '';
        this.elements.previewContainer.classList.add('hidden');
        this.elements.reportContainer.innerHTML = `<div class="flex flex-col items-center justify-center py-20 text-slate-400 bg-slate-50/30"><div class="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mb-4 text-slate-300"><i class="fas fa-table text-2xl"></i></div><p>Aguardando arquivo...</p></div>`;
        this.elements.processButton.disabled = true;
        this.elements.previewButton.disabled = true;
        ProgressManager.hide();
        const today = new Date();
        this.elements.periodoDe.value = new Date(today.getFullYear(), today.getMonth(), 1).toISOString().split('T')[0];
        this.elements.periodoAte.value = new Date(today.getFullYear(), today.getMonth() + 1, 0).toISOString().split('T')[0];
    }

    async showPreview() {
        if (!this.validateForm()) {
            NotificationManager.warning("Preencha todos os campos e carregue um arquivo para visualizar.");
            return;
        }
        ProgressManager.show();
        ProgressManager.update(20, 'Simulando dados...');
        try {
            const payload = {
                action: 'getPreviewHtml',
                spreadsheetId: AppState.selectedSpreadsheetId,
                medicoes: AppState.medicoes,
                logoContent: AppState.logoData,
                periodoDe: this.elements.periodoDe.value,
                periodoAte: this.elements.periodoAte.value,
                proximaLeitura: this.elements.proximaLeitura.value,
                tarifaEnergia: this.elements.tarifaEnergia.value.replace(',', '.'),
                taxaGestao: this.elements.taxaGestao.value.replace(',', '.'),
                rateioAreaComum: this.elements.rateioAreaComum.checked
            };
            const response = await fetch(APPS_SCRIPT_API_URL, {
                method: 'POST',
                body: JSON.stringify(payload),
                redirect: 'follow'
            });
            const result = await response.json();
            if (!result.success) throw new Error(result.message);
            AppState.previewData = result.previews;
            this.buildPreviewTabs();
            this.elements.previewContainer.classList.remove('hidden');
            this.elements.previewContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
        } catch (error) {
            NotificationManager.error(`Erro ao gerar simulação: ${error.message}`);
        } finally {
            ProgressManager.hide();
        }
    }

    buildPreviewTabs() {
        const tabsContainer = this.elements.previewTabs;
        if (!tabsContainer) return;
        tabsContainer.innerHTML = '';
        
        const globalTab = document.createElement('button');
        globalTab.className = 'preview-tab active';
        globalTab.innerHTML = '<i class="fas fa-globe mr-2"></i>Global';
        globalTab.dataset.target = 'global';
        tabsContainer.appendChild(globalTab);

        AppState.previewData.individuals.forEach((individual, index) => {
            const individualTab = document.createElement('button');
            individualTab.className = 'preview-tab';
            individualTab.textContent = `Unid. ${individual.unidade}`;
            individualTab.dataset.target = `individual-${index}`;
            tabsContainer.appendChild(individualTab);
        });
        
        const newTabsContainer = tabsContainer.cloneNode(true);
        tabsContainer.parentNode.replaceChild(newTabsContainer, tabsContainer);
        this.elements.previewTabs = newTabsContainer;

        newTabsContainer.addEventListener('click', (e) => {
            const targetButton = e.target.closest('.preview-tab');
            if (targetButton) {
                newTabsContainer.querySelectorAll('.preview-tab').forEach(tab => tab.classList.remove('active'));
                targetButton.classList.add('active');
                this.showPreviewTabContent(targetButton.dataset.target);
            }
        });

        this.showPreviewTabContent('global');
    }

    showPreviewTabContent(target) {
        const contentContainer = this.elements.previewTabContent;
        if (!contentContainer) return;

        let htmlContent = '';
        if (target === 'global') {
            htmlContent = AppState.previewData.global;
        } else if (target.startsWith('individual-')) {
            const index = parseInt(target.split('-')[1], 10);
            htmlContent = AppState.previewData.individuals[index].html;
        }

        if (!htmlContent) {
            contentContainer.innerHTML = '<div class="flex flex-col items-center justify-center h-64 text-slate-400"><i class="fas fa-eye-slash text-4xl mb-4"></i><p>Visualização indisponível</p></div>';
            return;
        }

        contentContainer.innerHTML = '';
        const iframe = document.createElement('iframe');
        iframe.className = 'w-full h-[600px] border-0 rounded-lg shadow-inner bg-white';
        iframe.title = `Pré-visualização do Relatório: ${target}`;
        
        contentContainer.appendChild(iframe);
        
        const iframeDoc = iframe.contentWindow.document;
        iframeDoc.open();
        iframeDoc.write(htmlContent);
        iframeDoc.close();
    }


    async handleProcessRequest() {
        if (!this.validateForm()) {
            NotificationManager.warning("Preencha todos os campos para continuar.");
            return;
        }
        ProgressManager.show();
        ProgressManager.update(10, 'Iniciando...');
        try {
            const payload = {
                action: 'processReport',
                spreadsheetId: AppState.selectedSpreadsheetId,
                medicoes: AppState.medicoes,
                logoContent: AppState.logoData,
                periodoDe: this.elements.periodoDe.value,
                periodoAte: this.elements.periodoAte.value,
                proximaLeitura: this.elements.proximaLeitura.value,
                tarifaEnergia: this.elements.tarifaEnergia.value.replace(',', '.'),
                taxaGestao: this.elements.taxaGestao.value.replace(',', '.'),
                rateioAreaComum: this.elements.rateioAreaComum.checked
            };
            ProgressManager.update(30, 'Enviando dados...');
            const response = await fetch(APPS_SCRIPT_API_URL, {
                method: 'POST',
                body: JSON.stringify(payload),
                redirect: 'follow'
            });
            ProgressManager.update(70, 'Gerando PDFs...');
            const result = await response.json();
            if (!result.success) throw new Error(result.message);
            ProgressManager.update(100, 'Finalizado!');
            NotificationManager.success(result.message);
            this.displayDownloadLinks(result.downloadLinks);
            
            setTimeout(() => {
                this.elements.downloadLinksSection.scrollIntoView({ behavior: 'smooth' });
            }, 500);

        } catch (error) {
            ProgressManager.hide();
            NotificationManager.error(`Erro: ${error.message}`);
        }
    }

    displayDownloadLinks(links) {
        this.elements.downloadLinksContainer.innerHTML = `
            <a href="${links.globalPdfUrl}" target="_blank" class="bg-tech-blue hover:bg-tech-dark text-white px-5 py-3 rounded font-bold shadow-md transition-all flex items-center transform hover:-translate-y-1">
                <i class="fas fa-file-pdf mr-2"></i> Relatório Global
            </a>
            <a href="${links.individualZipUrl}" target="_blank" class="bg-white border-2 border-slate-200 hover:border-tech-blue text-tech-blue hover:text-tech-dark px-5 py-3 rounded font-bold shadow-sm transition-all flex items-center transform hover:-translate-y-1">
                <i class="fas fa-file-archive mr-2"></i> Relatórios Individuais (ZIP)
            </a>
        `;
        this.elements.downloadLinksSection.classList.remove('hidden');
        ProgressManager.hide();
    }
}

document.addEventListener('DOMContentLoaded', () => new DashboardApp());