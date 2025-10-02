// SharePoint Versioning Manager - Web Edition JavaScript
// Versão: 2.0

class SharePointVersioningManager {
    constructor() {
        this.config = {
            majorVersions: 3,
            minorVersions: 1,
            tenantUrl: '',
            accessToken: null
        };
        this.sites = [];
        this.isProcessing = false;
        this.processResults = [];
        
        this.init();
    }

    init() {
        this.loadConfig();
        this.updateUI();
        this.bindEvents();
        this.addLog('INFO', 'Sistema iniciado com sucesso');
    }

    // Configuration Management
    loadConfig() {
        const saved = localStorage.getItem('spvm_config');
        if (saved) {
            this.config = { ...this.config, ...JSON.parse(saved) };
        }
        
        const savedSites = localStorage.getItem('spvm_sites');
        if (savedSites) {
            this.sites = JSON.parse(savedSites);
        }
    }

    saveConfig() {
        localStorage.setItem('spvm_config', JSON.stringify(this.config));
        localStorage.setItem('spvm_sites', JSON.stringify(this.sites));
        
        document.getElementById('majorVersions').value = this.config.majorVersions;
        document.getElementById('minorVersions').value = this.config.minorVersions;
        document.getElementById('tenantUrl').value = this.config.tenantUrl;
        
        this.updateProcessingInfo();
        this.addLog('SUCCESS', 'Configurações salvas com sucesso');
        this.showNotification('Configurações salvas!', 'success');
    }

    // UI Management
    updateUI() {
        document.getElementById('majorVersions').value = this.config.majorVersions;
        document.getElementById('minorVersions').value = this.config.minorVersions;
        document.getElementById('tenantUrl').value = this.config.tenantUrl;
        document.getElementById('sitesList').value = this.sites.join('\n');
        
        this.updateSitesCount();
        this.updateProcessingInfo();
        this.updateAuthStatus();
    }

    updateSitesCount() {
        const count = this.sites.length;
        document.getElementById('sitesCount').textContent = count;
        document.getElementById('processingSitesCount').textContent = count;
    }

    updateProcessingInfo() {
        document.getElementById('processingMajorVersions').textContent = this.config.majorVersions;
        document.getElementById('processingMinorVersions').textContent = this.config.minorVersions;
    }

    updateAuthStatus() {
        const statusElement = document.getElementById('authStatus');
        const buttonElement = document.getElementById('authButton');
        const connectionStatus = document.getElementById('connectionStatus');
        
        if (this.config.accessToken) {
            statusElement.className = 'auth-status authenticated';
            statusElement.innerHTML = '<i class="fas fa-check-circle text-success"></i><span>Autenticado com sucesso</span>';
            buttonElement.innerHTML = '<i class="fas fa-sign-out-alt"></i> Fazer Logout';
            buttonElement.className = 'btn btn-danger';
            connectionStatus.innerHTML = '<i class="fas fa-circle online"></i><span>Conectado</span>';
        } else {
            statusElement.className = 'auth-status not-authenticated';
            statusElement.innerHTML = '<i class="fas fa-exclamation-triangle text-warning"></i><span>Não autenticado</span>';
            buttonElement.innerHTML = '<i class="fas fa-sign-in-alt"></i> Fazer Login no SharePoint';
            buttonElement.className = 'btn btn-success';
            connectionStatus.innerHTML = '<i class="fas fa-circle offline"></i><span>Desconectado</span>';
        }
    }

    // Event Binding
    bindEvents() {
        // Config form changes
        document.getElementById('majorVersions').addEventListener('change', (e) => {
            this.config.majorVersions = parseInt(e.target.value);
        });
        
        document.getElementById('minorVersions').addEventListener('change', (e) => {
            this.config.minorVersions = parseInt(e.target.value);
        });
        
        document.getElementById('tenantUrl').addEventListener('change', (e) => {
            this.config.tenantUrl = e.target.value.trim();
        });

        // Sites list changes
        document.getElementById('sitesList').addEventListener('input', (e) => {
            const sites = e.target.value.split('\n')
                .map(s => s.trim())
                .filter(s => s.length > 0);
            this.sites = sites;
            this.updateSitesCount();
        });
    }

    // Tab Management
    showTab(tabName) {
        // Hide all tabs
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Remove active class from all nav tabs
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        
        // Show selected tab
        document.getElementById(tabName).classList.add('active');
        
        // Add active class to clicked nav tab
        event.target.classList.add('active');
        
        this.addLog('INFO', `Navegando para aba: ${tabName}`);
    }

    // Authentication
    async authenticate() {
        if (this.config.accessToken) {
            // Logout
            this.config.accessToken = null;
            this.saveConfig();
            this.updateAuthStatus();
            this.addLog('INFO', 'Logout realizado com sucesso');
            this.showNotification('Logout realizado!', 'info');
            return;
        }

        if (!this.config.tenantUrl) {
            this.showNotification('Configure a URL do tenant primeiro!', 'error');
            this.showTab('config');
            return;
        }

        try {
            this.addLog('INFO', 'Iniciando processo de autenticação...');
            this.showNotification('Abrindo janela de autenticação...', 'info');
            
            // Simular autenticação (em produção, usar MSAL.js)
            const result = await this.simulateAuthentication();
            
            if (result.success) {
                this.config.accessToken = result.token;
                this.saveConfig();
                this.updateAuthStatus();
                this.addLog('SUCCESS', 'Autenticação realizada com sucesso');
                this.showNotification('Autenticado com sucesso!', 'success');
            } else {
                throw new Error(result.error);
            }
            
        } catch (error) {
            this.addLog('ERROR', `Falha na autenticação: ${error.message}`);
            this.showNotification('Falha na autenticação!', 'error');
        }
    }

    async simulateAuthentication() {
        // Simular delay de autenticação
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Simular sucesso na autenticação (em produção, usar MSAL.js real)
        return {
            success: true,
            token: 'simulated_access_token_' + Date.now()
        };
    }

    // Sites Management
    loadSampleSites() {
        const sampleSites = [
            `${this.config.tenantUrl}/sites/exemplo-site-1`,
            `${this.config.tenantUrl}/sites/exemplo-site-2`,
            `${this.config.tenantUrl}/sites/exemplo-site-3`
        ];
        
        if (!this.config.tenantUrl) {
            this.showNotification('Configure a URL do tenant primeiro!', 'warning');
            return;
        }
        
        document.getElementById('sitesList').value = sampleSites.join('\n');
        this.sites = sampleSites;
        this.updateSitesCount();
        this.addLog('INFO', `Carregados ${sampleSites.length} sites de exemplo`);
        this.showNotification('Sites de exemplo carregados!', 'success');
    }

    validateSites() {
        const validationResults = document.getElementById('sitesValidation');
        
        if (this.sites.length === 0) {
            validationResults.className = 'validation-results error';
            validationResults.innerHTML = '<strong>Erro:</strong> Nenhum site na lista!';
            validationResults.style.display = 'block';
            return;
        }

        let validSites = 0;
        let invalidSites = [];

        this.sites.forEach(site => {
            if (this.isValidSharePointUrl(site)) {
                validSites++;
            } else {
                invalidSites.push(site);
            }
        });

        if (invalidSites.length === 0) {
            validationResults.className = 'validation-results success';
            validationResults.innerHTML = `<strong>Sucesso:</strong> Todos os ${validSites} sites são válidos!`;
            this.addLog('SUCCESS', `Validação concluída: ${validSites} sites válidos`);
        } else {
            validationResults.className = 'validation-results error';
            validationResults.innerHTML = `<strong>Atenção:</strong> ${invalidSites.length} URLs inválidas:<br>` +
                invalidSites.map(site => `• ${site}`).join('<br>');
            this.addLog('WARNING', `Validação encontrou ${invalidSites.length} URLs inválidas`);
        }

        validationResults.style.display = 'block';
    }

    isValidSharePointUrl(url) {
        const pattern = /^https:\/\/[a-zA-Z0-9-]+\.sharepoint\.com\/sites\/[a-zA-Z0-9-_]+\/?$/;
        return pattern.test(url.trim());
    }

    clearSites() {
        if (confirm('Tem certeza que deseja limpar toda a lista de sites?')) {
            document.getElementById('sitesList').value = '';
            this.sites = [];
            this.updateSitesCount();
            document.getElementById('sitesValidation').style.display = 'none';
            this.addLog('INFO', 'Lista de sites limpa');
            this.showNotification('Lista limpa!', 'info');
        }
    }

    exportSites() {
        if (this.sites.length === 0) {
            this.showNotification('Nenhum site para exportar!', 'warning');
            return;
        }

        const data = this.sites.join('\n');
        const blob = new Blob([data], { type: 'text/plain' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'sharepoint-sites-list.txt';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        this.addLog('INFO', `Lista de ${this.sites.length} sites exportada`);
        this.showNotification('Lista exportada!', 'success');
    }

    // Processing
    async startProcessing() {
        if (!this.config.accessToken) {
            this.showNotification('Faça login primeiro!', 'error');
            this.showTab('config');
            return;
        }

        if (this.sites.length === 0) {
            this.showNotification('Adicione sites à lista primeiro!', 'error');
            this.showTab('sites');
            return;
        }

        if (this.isProcessing) {
            this.showNotification('Processamento já em andamento!', 'warning');
            return;
        }

        this.isProcessing = true;
        this.processResults = [];
        
        // Update UI
        document.getElementById('startProcessButton').style.display = 'none';
        document.getElementById('stopProcessButton').style.display = 'inline-flex';
        document.getElementById('progressSection').style.display = 'block';
        document.getElementById('resultsSection').style.display = 'none';
        
        this.addLog('INFO', `Iniciando processamento de ${this.sites.length} sites`);
        this.addLog('WARNING', 'IMPORTANTE: Mantenha-se próximo para autenticar cada site!');
        
        try {
            for (let i = 0; i < this.sites.length && this.isProcessing; i++) {
                const site = this.sites[i];
                const progress = ((i + 1) / this.sites.length) * 100;
                
                this.updateProgress(progress, `Processando: ${site}`);
                this.addLog('INFO', `[${i + 1}/${this.sites.length}] Processando site: ${site}`);
                
                const result = await this.processSite(site);
                this.processResults.push(result);
                
                if (result.success) {
                    this.addLog('SUCCESS', `Site processado: ${result.librariesProcessed} bibliotecas configuradas`);
                } else {
                    this.addLog('ERROR', `Falha no site: ${result.error}`);
                }
                
                // Pausa entre sites
                if (i < this.sites.length - 1 && this.isProcessing) {
                    this.addLog('INFO', 'Aguardando 2 segundos antes do próximo site...');
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }
            
            if (this.isProcessing) {
                this.completeProcessing();
            }
            
        } catch (error) {
            this.addLog('ERROR', `Erro durante processamento: ${error.message}`);
            this.showNotification('Erro durante processamento!', 'error');
        } finally {
            this.isProcessing = false;
            document.getElementById('startProcessButton').style.display = 'inline-flex';
            document.getElementById('stopProcessButton').style.display = 'none';
        }
    }

    async processSite(siteUrl) {
        try {
            this.addLog('INFO', `Conectando ao site: ${siteUrl}`);
            this.addLog('WARNING', 'Janela de autenticação será aberta...');
            
            // Simular conexão e processamento
            await new Promise(resolve => setTimeout(resolve, 1500));
            
            // Simular obtenção de bibliotecas
            const librariesCount = Math.floor(Math.random() * 5) + 1; // 1-5 bibliotecas
            this.addLog('INFO', `Encontradas ${librariesCount} bibliotecas de documentos`);
            
            // Simular configuração de cada biblioteca
            let processedLibraries = 0;
            for (let i = 0; i < librariesCount; i++) {
                await new Promise(resolve => setTimeout(resolve, 500));
                
                // 90% de chance de sucesso
                if (Math.random() > 0.1) {
                    processedLibraries++;
                    this.addLog('SUCCESS', `Biblioteca 'Documentos ${i + 1}' configurada`);
                } else {
                    this.addLog('ERROR', `Falha na biblioteca 'Documentos ${i + 1}'`);
                }
            }
            
            this.addLog('INFO', `Desconectado do site: ${siteUrl}`);
            
            return {
                siteUrl: siteUrl,
                success: processedLibraries > 0,
                librariesProcessed: processedLibraries,
                librariesFailed: librariesCount - processedLibraries,
                totalLibraries: librariesCount,
                error: processedLibraries === 0 ? 'Nenhuma biblioteca foi processada' : null
            };
            
        } catch (error) {
            return {
                siteUrl: siteUrl,
                success: false,
                librariesProcessed: 0,
                librariesFailed: 0,
                totalLibraries: 0,
                error: error.message
            };
        }
    }

    updateProgress(percentage, currentSite) {
        document.getElementById('progressFill').style.width = `${percentage}%`;
        document.getElementById('progressText').textContent = `${Math.round(percentage)}%`;
        document.getElementById('currentSite').textContent = currentSite;
    }

    completeProcessing() {
        this.updateProgress(100, 'Processamento concluído!');
        
        // Calculate results
        const successCount = this.processResults.filter(r => r.success).length;
        const errorCount = this.processResults.length - successCount;
        const totalLibraries = this.processResults.reduce((sum, r) => sum + r.librariesProcessed, 0);
        
        // Update results display
        document.getElementById('successCount').textContent = successCount;
        document.getElementById('errorCount').textContent = errorCount;
        document.getElementById('totalLibraries').textContent = totalLibraries;
        document.getElementById('resultsSection').style.display = 'block';
        
        // Log final results
        this.addLog('INFO', '=== RELATÓRIO FINAL ===');
        this.addLog('SUCCESS', `Sites processados com sucesso: ${successCount}`);
        this.addLog('ERROR', `Sites com falha: ${errorCount}`);
        this.addLog('SUCCESS', `Total de bibliotecas configuradas: ${totalLibraries}`);
        
        this.showNotification('Processamento concluído!', 'success');
        
        // Save results for reports
        this.saveProcessingReport();
    }

    stopProcessing() {
        if (confirm('Tem certeza que deseja parar o processamento?')) {
            this.isProcessing = false;
            this.addLog('WARNING', 'Processamento interrompido pelo usuário');
            this.showNotification('Processamento interrompido!', 'warning');
        }
    }

    saveProcessingReport() {
        const report = {
            timestamp: new Date().toISOString(),
            config: { ...this.config },
            results: this.processResults,
            summary: {
                totalSites: this.processResults.length,
                successfulSites: this.processResults.filter(r => r.success).length,
                failedSites: this.processResults.filter(r => !r.success).length,
                totalLibraries: this.processResults.reduce((sum, r) => sum + r.librariesProcessed, 0)
            }
        };
        
        const reports = JSON.parse(localStorage.getItem('spvm_reports') || '[]');
        reports.unshift(report); // Add to beginning
        
        // Keep only last 10 reports
        if (reports.length > 10) {
            reports.splice(10);
        }
        
        localStorage.setItem('spvm_reports', JSON.stringify(reports));
        this.loadReports();
    }

    loadReports() {
        const reports = JSON.parse(localStorage.getItem('spvm_reports') || '[]');
        const reportsList = document.getElementById('reportsList');
        
        if (reports.length === 0) {
            reportsList.innerHTML = '<p class="no-reports">Nenhum relatório disponível. Execute um processamento primeiro.</p>';
            return;
        }
        
        reportsList.innerHTML = reports.map((report, index) => `
            <div class="card">
                <div class="card-header">
                    <h3><i class="fas fa-file-alt"></i> Relatório ${index + 1}</h3>
                    <span>${new Date(report.timestamp).toLocaleString('pt-BR')}</span>
                </div>
                <div class="card-body">
                    <div class="results-grid">
                        <div class="result-item success">
                            <div class="result-number">${report.summary.successfulSites}</div>
                            <div class="result-label">Sites OK</div>
                        </div>
                        <div class="result-item error">
                            <div class="result-number">${report.summary.failedSites}</div>
                            <div class="result-label">Sites Erro</div>
                        </div>
                        <div class="result-item total">
                            <div class="result-number">${report.summary.totalLibraries}</div>
                            <div class="result-label">Bibliotecas</div>
                        </div>
                    </div>
                    <div class="mt-20">
                        <button class="btn btn-info btn-small" onclick="app.showReportDetails(${index})">
                            <i class="fas fa-eye"></i> Ver Detalhes
                        </button>
                        <button class="btn btn-secondary btn-small" onclick="app.exportReport(${index})">
                            <i class="fas fa-download"></i> Exportar
                        </button>
                    </div>
                </div>
            </div>
        `).join('');
    }

    showReportDetails(index) {
        const reports = JSON.parse(localStorage.getItem('spvm_reports') || '[]');
        const report = reports[index];
        
        if (!report) return;
        
        const details = report.results.map(result => `
            <tr>
                <td>${result.siteUrl}</td>
                <td class="${result.success ? 'text-success' : 'text-error'}">
                    ${result.success ? '✅ Sucesso' : '❌ Falha'}
                </td>
                <td>${result.librariesProcessed}</td>
                <td>${result.librariesFailed}</td>
                <td>${result.error || '-'}</td>
            </tr>
        `).join('');
        
        const popup = window.open('', '_blank', 'width=800,height=600');
        popup.document.write(`
            <html>
                <head>
                    <title>Detalhes do Relatório</title>
                    <style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                        th { background-color: #f2f2f2; }
                        .text-success { color: #107c10; }
                        .text-error { color: #d83b01; }
                    </style>
                </head>
                <body>
                    <h1>Relatório Detalhado</h1>
                    <p><strong>Data:</strong> ${new Date(report.timestamp).toLocaleString('pt-BR')}</p>
                    <p><strong>Configuração:</strong> ${report.config.majorVersions} versões principais, ${report.config.minorVersions} versões secundárias</p>
                    
                    <table>
                        <thead>
                            <tr>
                                <th>Site URL</th>
                                <th>Status</th>
                                <th>Bibliotecas OK</th>
                                <th>Bibliotecas Erro</th>
                                <th>Erro</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${details}
                        </tbody>
                    </table>
                </body>
            </html>
        `);
    }

    exportReport(index) {
        const reports = JSON.parse(localStorage.getItem('spvm_reports') || '[]');
        const report = reports[index];
        
        if (!report) return;
        
        const csvContent = [
            'Site URL,Status,Bibliotecas OK,Bibliotecas Erro,Erro',
            ...report.results.map(result => 
                `"${result.siteUrl}","${result.success ? 'Sucesso' : 'Falha'}",${result.librariesProcessed},${result.librariesFailed},"${result.error || ''}"`
            )
        ].join('\n');
        
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `sharepoint-report-${new Date(report.timestamp).toISOString().split('T')[0]}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        this.showNotification('Relatório exportado!', 'success');
    }

    // Logging
    addLog(level, message) {
        const logContainer = document.getElementById('liveLog');
        const timestamp = new Date().toLocaleTimeString('pt-BR');
        
        const logEntry = document.createElement('div');
        logEntry.className = 'log-entry';
        logEntry.innerHTML = `
            <span class="log-timestamp">[${timestamp}]</span>
            <span class="log-level ${level}">${level}</span>
            <span class="log-message">${message}</span>
        `;
        
        logContainer.appendChild(logEntry);
        logContainer.scrollTop = logContainer.scrollHeight;
        
        // Keep only last 100 log entries
        while (logContainer.children.length > 100) {
            logContainer.removeChild(logContainer.firstChild);
        }
    }

    clearLog() {
        document.getElementById('liveLog').innerHTML = '';
        this.addLog('INFO', 'Log limpo');
    }

    // Notifications
    showNotification(message, type = 'info') {
        // Create notification element
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 15px 20px;
            border-radius: 6px;
            color: white;
            font-weight: 600;
            z-index: 1000;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            transform: translateX(100%);
            transition: transform 0.3s ease;
        `;
        
        // Set background color based on type
        const colors = {
            success: '#107c10',
            error: '#d83b01',
            warning: '#ff8c00',
            info: '#0078d4'
        };
        notification.style.backgroundColor = colors[type] || colors.info;
        notification.textContent = message;
        
        document.body.appendChild(notification);
        
        // Animate in
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);
        
        // Remove after 3 seconds
        setTimeout(() => {
            notification.style.transform = 'translateX(100%)';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 3000);
    }
}

// Global functions for HTML onclick events
function showTab(tabName) {
    app.showTab(tabName);
}

function saveConfig() {
    app.config.majorVersions = parseInt(document.getElementById('majorVersions').value);
    app.config.minorVersions = parseInt(document.getElementById('minorVersions').value);
    app.config.tenantUrl = document.getElementById('tenantUrl').value.trim();
    app.saveConfig();
}

function authenticate() {
    app.authenticate();
}

function loadSampleSites() {
    app.loadSampleSites();
}

function validateSites() {
    app.validateSites();
}

function clearSites() {
    app.clearSites();
}

function exportSites() {
    app.exportSites();
}

function startProcessing() {
    app.startProcessing();
}

function stopProcessing() {
    app.stopProcessing();
}

function clearLog() {
    app.clearLog();
}

// Initialize application when DOM is loaded
let app;
document.addEventListener('DOMContentLoaded', function() {
    app = new SharePointVersioningManager();
    app.loadReports();
});

// Handle tab navigation with keyboard
document.addEventListener('keydown', function(e) {
    if (e.ctrlKey) {
        const tabMap = {
            '1': 'config',
            '2': 'sites', 
            '3': 'process',
            '4': 'reports',
            '5': 'help'
        };
        
        if (tabMap[e.key]) {
            e.preventDefault();
            app.showTab(tabMap[e.key]);
        }
    }
});
