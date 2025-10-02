// SharePoint Versioning Manager - Web Edition com AUTENTICAÇÃO REAL
// Versão: 2.1 - Microsoft Authentication Library (MSAL.js)

class SharePointVersioningManagerReal {
    constructor() {
        this.config = {
            majorVersions: 3,
            minorVersions: 1,
            tenantUrl: '',
            accessToken: null,
            userInfo: null
        };
        this.sites = [];
        this.isProcessing = false;
        this.processResults = [];
        
        // Configuração MSAL para autenticação real
        this.msalConfig = {
            auth: {
                clientId: "14d82eec-204b-4c2f-b7e0-446a3b5b2faa", // App ID público da Microsoft
                authority: "https://login.microsoftonline.com/common",
                redirectUri: window.location.origin
            },
            cache: {
                cacheLocation: "localStorage",
                storeAuthStateInCookie: false
            }
        };
        
        // Escopos necessários para SharePoint
        this.loginRequest = {
            scopes: [
                "https://graph.microsoft.com/Sites.ReadWrite.All",
                "https://graph.microsoft.com/User.Read",
                "openid",
                "profile"
            ]
        };
        
        this.init();
    }

    async init() {
        this.loadConfig();
        await this.initializeMSAL();
        this.updateUI();
        this.bindEvents();
        this.addLog('INFO', 'Sistema iniciado com autenticação real Microsoft');
    }

    // Inicializar MSAL (Microsoft Authentication Library)
    async initializeMSAL() {
        try {
            // Carregar MSAL.js dinamicamente
            await this.loadMSALLibrary();
            
            // Inicializar instância MSAL
            this.msalInstance = new msal.PublicClientApplication(this.msalConfig);
            
            // Verificar se já existe uma conta logada
            const accounts = this.msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                this.config.userInfo = accounts[0];
                await this.acquireTokenSilent();
            }
            
            this.addLog('SUCCESS', 'MSAL inicializado com sucesso');
        } catch (error) {
            this.addLog('ERROR', `Erro ao inicializar MSAL: ${error.message}`);
            this.addLog('WARNING', 'Usando modo de demonstração');
        }
    }

    // Carregar biblioteca MSAL.js
    loadMSALLibrary() {
        return new Promise((resolve, reject) => {
            if (window.msal) {
                resolve();
                return;
            }

            const script = document.createElement('script');
            script.src = 'https://alcdn.msauth.net/browser/2.38.4/js/msal-browser.min.js';
            script.onload = () => {
                this.addLog('SUCCESS', 'Biblioteca MSAL carregada');
                resolve();
            };
            script.onerror = () => {
                this.addLog('ERROR', 'Falha ao carregar biblioteca MSAL');
                reject(new Error('Falha ao carregar MSAL'));
            };
            document.head.appendChild(script);
        });
    }

    // Autenticação real com Microsoft
    async authenticate() {
        if (this.config.accessToken && this.config.userInfo) {
            // Logout
            await this.logout();
            return;
        }

        if (!this.config.tenantUrl) {
            this.showNotification('Configure a URL do tenant primeiro!', 'error');
            this.showTab('config');
            return;
        }

        try {
            this.addLog('INFO', 'Iniciando autenticação Microsoft...');
            this.showNotification('Abrindo janela de login Microsoft...', 'info');
            
            if (!this.msalInstance) {
                throw new Error('MSAL não inicializado');
            }

            // Fazer login com popup
            const loginResponse = await this.msalInstance.loginPopup(this.loginRequest);
            
            this.config.userInfo = loginResponse.account;
            this.config.accessToken = loginResponse.accessToken;
            
            this.saveConfig();
            this.updateAuthStatus();
            
            this.addLog('SUCCESS', `Autenticado como: ${this.config.userInfo.name} (${this.config.userInfo.username})`);
            this.showNotification(`Bem-vindo, ${this.config.userInfo.name}!`, 'success');
            
        } catch (error) {
            this.addLog('ERROR', `Falha na autenticação: ${error.message}`);
            
            if (error.message.includes('popup_window_error')) {
                this.addLog('WARNING', 'Popup bloqueado. Permitir popups e tentar novamente.');
                this.showNotification('Popup bloqueado! Permita popups e tente novamente.', 'warning');
            } else {
                this.showNotification('Falha na autenticação!', 'error');
            }
        }
    }

    // Obter token silenciosamente (renovar se necessário)
    async acquireTokenSilent() {
        try {
            const tokenRequest = {
                ...this.loginRequest,
                account: this.config.userInfo
            };
            
            const tokenResponse = await this.msalInstance.acquireTokenSilent(tokenRequest);
            this.config.accessToken = tokenResponse.accessToken;
            
            this.addLog('SUCCESS', 'Token renovado automaticamente');
            return tokenResponse.accessToken;
            
        } catch (error) {
            this.addLog('WARNING', 'Token expirado, necessário fazer login novamente');
            await this.logout();
            throw error;
        }
    }

    // Logout
    async logout() {
        try {
            if (this.msalInstance && this.config.userInfo) {
                await this.msalInstance.logoutPopup({
                    account: this.config.userInfo
                });
            }
            
            this.config.accessToken = null;
            this.config.userInfo = null;
            this.saveConfig();
            this.updateAuthStatus();
            
            this.addLog('INFO', 'Logout realizado com sucesso');
            this.showNotification('Logout realizado!', 'info');
            
        } catch (error) {
            this.addLog('ERROR', `Erro no logout: ${error.message}`);
            // Forçar logout local mesmo se der erro
            this.config.accessToken = null;
            this.config.userInfo = null;
            this.saveConfig();
            this.updateAuthStatus();
        }
    }

    // Atualizar status de autenticação
    updateAuthStatus() {
        const statusElement = document.getElementById('authStatus');
        const buttonElement = document.getElementById('authButton');
        const connectionStatus = document.getElementById('connectionStatus');
        
        if (this.config.accessToken && this.config.userInfo) {
            statusElement.className = 'auth-status authenticated';
            statusElement.innerHTML = `
                <i class="fas fa-check-circle text-success"></i>
                <div>
                    <strong>Autenticado como:</strong><br>
                    <span>${this.config.userInfo.name}</span><br>
                    <small>${this.config.userInfo.username}</small>
                </div>
            `;
            buttonElement.innerHTML = '<i class="fas fa-sign-out-alt"></i> Fazer Logout';
            buttonElement.className = 'btn btn-danger';
            connectionStatus.innerHTML = '<i class="fas fa-circle online"></i><span>Conectado</span>';
        } else {
            statusElement.className = 'auth-status not-authenticated';
            statusElement.innerHTML = '<i class="fas fa-exclamation-triangle text-warning"></i><span>Não autenticado</span>';
            buttonElement.innerHTML = '<i class="fas fa-sign-in-alt"></i> Login Microsoft';
            buttonElement.className = 'btn btn-success';
            connectionStatus.innerHTML = '<i class="fas fa-circle offline"></i><span>Desconectado</span>';
        }
    }

    // Processar site com autenticação real
    async processSite(siteUrl) {
        try {
            this.addLog('INFO', `Conectando ao site: ${siteUrl}`);
            
            // Verificar se token ainda é válido
            if (!this.config.accessToken) {
                throw new Error('Token de acesso não disponível');
            }

            // Tentar renovar token se necessário
            try {
                await this.acquireTokenSilent();
            } catch (tokenError) {
                throw new Error('Sessão expirada, faça login novamente');
            }

            // Extrair informações do site URL
            const urlParts = siteUrl.split('/');
            const tenantName = urlParts[2].split('.')[0];
            const siteName = urlParts[urlParts.length - 1];

            // Fazer chamada real para SharePoint REST API
            const siteResponse = await this.callSharePointAPI(siteUrl, '/_api/web', 'GET');
            
            if (!siteResponse.ok) {
                throw new Error(`Erro ao acessar site: ${siteResponse.status}`);
            }

            const siteData = await siteResponse.json();
            this.addLog('SUCCESS', `Site conectado: ${siteData.Title}`);

            // Obter bibliotecas de documentos
            const listsResponse = await this.callSharePointAPI(
                siteUrl, 
                "/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false", 
                'GET'
            );

            if (!listsResponse.ok) {
                throw new Error(`Erro ao obter bibliotecas: ${listsResponse.status}`);
            }

            const listsData = await listsResponse.json();
            const libraries = listsData.value;

            this.addLog('INFO', `Encontradas ${libraries.length} bibliotecas de documentos`);

            let processedLibraries = 0;
            let failedLibraries = 0;

            // Configurar cada biblioteca
            for (const library of libraries) {
                try {
                    await this.configureLibraryVersioning(siteUrl, library);
                    processedLibraries++;
                    this.addLog('SUCCESS', `Biblioteca '${library.Title}' configurada`);
                } catch (libError) {
                    failedLibraries++;
                    this.addLog('ERROR', `Falha na biblioteca '${library.Title}': ${libError.message}`);
                }
            }

            this.addLog('INFO', `Processamento concluído para: ${siteUrl}`);

            return {
                siteUrl: siteUrl,
                success: processedLibraries > 0,
                librariesProcessed: processedLibraries,
                librariesFailed: failedLibraries,
                totalLibraries: libraries.length,
                error: processedLibraries === 0 ? 'Nenhuma biblioteca foi processada' : null
            };

        } catch (error) {
            this.addLog('ERROR', `Erro no site ${siteUrl}: ${error.message}`);
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

    // Configurar versionamento de uma biblioteca
    async configureLibraryVersioning(siteUrl, library) {
        const updateData = {
            EnableVersioning: true,
            MajorVersionLimit: this.config.majorVersions,
            MajorWithMinorVersionsLimit: this.config.minorVersions
        };

        const response = await this.callSharePointAPI(
            siteUrl,
            `/_api/web/lists(guid'${library.Id}')`,
            'PATCH',
            updateData
        );

        if (!response.ok) {
            throw new Error(`Falha ao configurar biblioteca: ${response.status}`);
        }
    }

    // Fazer chamada para SharePoint REST API
    async callSharePointAPI(siteUrl, endpoint, method = 'GET', data = null) {
        const url = `${siteUrl}${endpoint}`;
        
        const headers = {
            'Authorization': `Bearer ${this.config.accessToken}`,
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        };

        if (method === 'PATCH') {
            headers['X-HTTP-Method'] = 'MERGE';
            headers['If-Match'] = '*';
        }

        const options = {
            method: method,
            headers: headers
        };

        if (data && (method === 'POST' || method === 'PATCH')) {
            options.body = JSON.stringify(data);
        }

        return fetch(url, options);
    }

    // Herdar todas as outras funções da classe original
    loadConfig() {
        const saved = localStorage.getItem('spvm_config_real');
        if (saved) {
            const savedConfig = JSON.parse(saved);
            this.config = { ...this.config, ...savedConfig };
        }
        
        const savedSites = localStorage.getItem('spvm_sites');
        if (savedSites) {
            this.sites = JSON.parse(savedSites);
        }
    }

    saveConfig() {
        // Não salvar o token por segurança, apenas info do usuário
        const configToSave = {
            majorVersions: this.config.majorVersions,
            minorVersions: this.config.minorVersions,
            tenantUrl: this.config.tenantUrl,
            userInfo: this.config.userInfo
        };
        
        localStorage.setItem('spvm_config_real', JSON.stringify(configToSave));
        localStorage.setItem('spvm_sites', JSON.stringify(this.sites));
        
        document.getElementById('majorVersions').value = this.config.majorVersions;
        document.getElementById('minorVersions').value = this.config.minorVersions;
        document.getElementById('tenantUrl').value = this.config.tenantUrl;
        
        this.updateProcessingInfo();
        this.addLog('SUCCESS', 'Configurações salvas com sucesso');
        this.showNotification('Configurações salvas!', 'success');
    }

    // Resto das funções permanecem iguais à classe original...
    // (copiando as funções essenciais)

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

    bindEvents() {
        document.getElementById('majorVersions').addEventListener('change', (e) => {
            this.config.majorVersions = parseInt(e.target.value);
        });
        
        document.getElementById('minorVersions').addEventListener('change', (e) => {
            this.config.minorVersions = parseInt(e.target.value);
        });
        
        document.getElementById('tenantUrl').addEventListener('change', (e) => {
            this.config.tenantUrl = e.target.value.trim();
        });

        document.getElementById('sitesList').addEventListener('input', (e) => {
            const sites = e.target.value.split('\n')
                .map(s => s.trim())
                .filter(s => s.length > 0);
            this.sites = sites;
            this.updateSitesCount();
        });
    }

    showTab(tabName) {
        document.querySelectorAll('.tab-content').forEach(tab => {
            tab.classList.remove('active');
        });
        
        document.querySelectorAll('.nav-tab').forEach(tab => {
            tab.classList.remove('active');
        });
        
        document.getElementById(tabName).classList.add('active');
        event.target.classList.add('active');
        
        this.addLog('INFO', `Navegando para aba: ${tabName}`);
    }

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
        
        while (logContainer.children.length > 100) {
            logContainer.removeChild(logContainer.firstChild);
        }
    }

    showNotification(message, type = 'info') {
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
        
        const colors = {
            success: '#107c10',
            error: '#d83b01',
            warning: '#ff8c00',
            info: '#0078d4'
        };
        notification.style.backgroundColor = colors[type] || colors.info;
        notification.textContent = message;
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            notification.style.transform = 'translateX(0)';
        }, 100);
        
        setTimeout(() => {
            notification.style.transform = 'translateX(100%)';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 3000);
    }

    // Adicionar funções de processamento simplificadas
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

        this.isProcessing = true;
        this.processResults = [];
        
        document.getElementById('startProcessButton').style.display = 'none';
        document.getElementById('stopProcessButton').style.display = 'inline-flex';
        document.getElementById('progressSection').style.display = 'block';
        
        this.addLog('INFO', `Iniciando processamento REAL de ${this.sites.length} sites`);
        this.addLog('SUCCESS', `Usuário autenticado: ${this.config.userInfo.name}`);
        
        try {
            for (let i = 0; i < this.sites.length && this.isProcessing; i++) {
                const site = this.sites[i];
                const progress = ((i + 1) / this.sites.length) * 100;
                
                this.updateProgress(progress, `Processando: ${site}`);
                
                const result = await this.processSite(site);
                this.processResults.push(result);
                
                if (i < this.sites.length - 1 && this.isProcessing) {
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }
            
            if (this.isProcessing) {
                this.completeProcessing();
            }
            
        } catch (error) {
            this.addLog('ERROR', `Erro durante processamento: ${error.message}`);
        } finally {
            this.isProcessing = false;
            document.getElementById('startProcessButton').style.display = 'inline-flex';
            document.getElementById('stopProcessButton').style.display = 'none';
        }
    }

    updateProgress(percentage, currentSite) {
        document.getElementById('progressFill').style.width = `${percentage}%`;
        document.getElementById('progressText').textContent = `${Math.round(percentage)}%`;
        document.getElementById('currentSite').textContent = currentSite;
    }

    completeProcessing() {
        this.updateProgress(100, 'Processamento concluído!');
        
        const successCount = this.processResults.filter(r => r.success).length;
        const errorCount = this.processResults.length - successCount;
        const totalLibraries = this.processResults.reduce((sum, r) => sum + r.librariesProcessed, 0);
        
        document.getElementById('successCount').textContent = successCount;
        document.getElementById('errorCount').textContent = errorCount;
        document.getElementById('totalLibraries').textContent = totalLibraries;
        document.getElementById('resultsSection').style.display = 'block';
        
        this.addLog('INFO', '=== RELATÓRIO FINAL ===');
        this.addLog('SUCCESS', `Sites processados com sucesso: ${successCount}`);
        this.addLog('ERROR', `Sites com falha: ${errorCount}`);
        this.addLog('SUCCESS', `Total de bibliotecas configuradas: ${totalLibraries}`);
        
        this.showNotification('Processamento concluído!', 'success');
    }
}

// Funções globais para HTML
function showTab(tabName) {
    realApp.showTab(tabName);
}

function saveConfig() {
    realApp.config.majorVersions = parseInt(document.getElementById('majorVersions').value);
    realApp.config.minorVersions = parseInt(document.getElementById('minorVersions').value);
    realApp.config.tenantUrl = document.getElementById('tenantUrl').value.trim();
    realApp.saveConfig();
}

function authenticate() {
    realApp.authenticate();
}

function startProcessing() {
    realApp.startProcessing();
}

// Inicializar aplicação com autenticação real
let realApp;
document.addEventListener('DOMContentLoaded', function() {
    realApp = new SharePointVersioningManagerReal();
});
