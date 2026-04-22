// Importações do Firebase SDK
        import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
        import { getFirestore, collection, doc, addDoc, getDocs, setDoc, updateDoc, deleteDoc, onSnapshot, serverTimestamp, runTransaction, writeBatch, Timestamp, getDoc } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

        // --- Configuração do Firebase ---
        const firebaseConfig = {
            apiKey: "AIzaSyBs8HSita5HgY4mOu7EXGlcQsmbcuVUCEA",
            authDomain: "ferramentaria-com-facial-e6a38.firebaseapp.com",
            projectId: "ferramentaria-com-facial-e6a38",
            storageBucket: "ferramentaria-com-facial-e6a38.firebasestorage.app",
            messagingSenderId: "344365806314",
            appId: "1:344365806314:web:db0c3c2bae4b0bbf165d93"
        };

        const appId = 'ferramentaria-estoque'; // ID único para organizar os dados no Firestore
        const app = initializeApp(firebaseConfig);
        const db = getFirestore(app);
        // Sem Firebase Auth — autenticação gerenciada pelo próprio app

        // --- Estado da Aplicação ---
        let products = [];
        let history = [];
        let requisitions = [];
        let locations = [];
        let appSettings = { appName: 'UHE Estrela', logoUrl: null };
        let currentProductId = null;
        let currentLocationId = null;
        let productsCollectionRef;
        let historyCollectionRef;
        let requisitionsCollectionRef;
        let locationsCollectionRef;
        let settingsDocRef;
        let usersCollectionRef;
        let inventoryFilter = 'all';
        let inventorySortOrder = 'default';
        let isDataLoaded = false;
        let currentAudio = null;
        let selectedProductIds = new Set();
        let currentViewId = 'dashboard-view';
        let coreUnsubscribers = [];
        let estrelaUnsubscribers = [];
        let hasVisibilityListener = false;
        
        // 🔐 Sistema de Autenticação e Permissões
        let currentUser = null;
        let userRole = null; // 'admin', 'operador', 'visualizador'
        let isAuthInitialized = false; // Flag para controlar inicialização
        
        // 🚀 Paginação para Performance
        let currentPage = 1;
        const itemsPerPage = 50; // Mostrar 50 produtos por vez 
        let hasAutoFilledMissingGroups = false;
        let isAutoFillingMissingGroups = false;
        let dashboardTopItemsChart = null;
        let dashboardTrendChart = null;
        let lastComprasFilteredItems = [];
        let lastComprasPeriodDays = 15;
        
        // ⭐ UHE Estrela State
        let estrelaProducts = [];
        let estrelaEntries = [];
        let estrelaExits = [];
        let estrelaProductsRef;
        let estrelaEntriesRef;
        let estrelaExitsRef;

        // --- Seletores ---
        const productList = document.getElementById('product-list');
        const searchInput = document.getElementById('search-input');
        const noProductsMessage = document.getElementById('no-products-message');
        const authStatusDiv = document.getElementById('auth-status');
        const inventoryFilters = document.getElementById('inventory-filters');
        const dashboardStats = document.getElementById('dashboard-stats');
        const dashboardActivity = document.getElementById('dashboard-activity');
        const tabButtons = document.querySelectorAll('.tab-btn');
        const views = document.querySelectorAll('.view-content');
        const exitsList = document.getElementById('exits-list');
        const exitsSearchInput = document.getElementById('exits-search-input');
        const noExitsMessage = document.getElementById('no-exits-message');
        const sortOrderSelect = document.getElementById('sort-order');
        const addForm = document.getElementById('add-product-form');
        const importBtn = document.getElementById('import-btn');
        const exportBtn = document.getElementById('export-btn');
        const csvFileInput = document.getElementById('csv-file-input');
        const entryForm = document.getElementById('entry-form');
        const entryProductSelect = document.getElementById('entry-product-select') || document.createElement('input');
        const entryProductSearch = document.getElementById('entry-product-search');
        const entryProductsContainer = document.getElementById('entry-products-container');
        const loaderOverlay = document.getElementById('loader-overlay');
        const settingsBtn = document.getElementById('settings-btn');
        const createRequisitionBtn = document.getElementById('create-requisition-btn');
        const requisitionsList = document.getElementById('requisitions-list');
        const noRequisitionsMessage = document.getElementById('no-requisitions-message');
        const exitForm = document.getElementById('exit-form');
        const micBtn = document.getElementById('mic-btn');
        const printSelectedReqsBtn = document.getElementById('print-selected-reqs-btn');
        const generateExitReportBtn = document.getElementById('generate-exit-report-btn');
        const generateConsumptionReportBtn = document.getElementById('generate-consumption-report-btn');
        const generateKpiReportBtn = document.getElementById('generate-kpi-report-btn');
        const addLocationForm = document.getElementById('add-location-form');
        const loginUsernameInput = document.getElementById('login-username');
        const loginPasswordInput = document.getElementById('login-password');
        const loginBtn = document.getElementById('login-btn');
        const loginError = document.getElementById('login-error');
        const logoutBtn = document.getElementById('logout-btn');
        const loginScreen = document.getElementById('login-screen');
        const appContainer = document.getElementById('app-container');
        const userInfoDiv = document.getElementById('user-info');
        const locationsList = document.getElementById('locations-list');
        const noLocationsMessage = document.getElementById('no-locations-message');
        const deleteSelectedBtn = document.getElementById('delete-selected-btn');
        const selectAllProductsCheckbox = document.getElementById('select-all-products');
        const mobileMenuBtn = document.getElementById('mobile-menu-btn');
        const mainNav = document.getElementById('main-nav');
        const currentViewTitle = document.getElementById('current-view-title');
        const aiDescribeBtn = document.getElementById('ai-describe-btn');
        const backupBtn = document.getElementById('backup-btn');
        const restoreBtn = document.getElementById('restore-btn');
        const restoreFileInput = document.getElementById('restore-file-input');
        const initiateRequisitionBtn = document.getElementById('initiate-requisition-btn');
        const rmList = document.getElementById('rm-list');
        const rmSearchInput = document.getElementById('rm-search-input');
        const noRmMessage = document.getElementById('no-rm-message');
        const importLocationsBtn = document.getElementById('import-locations-btn');
        const csvLocationsFileInput = document.getElementById('csv-locations-file-input');
        const aiExitPrompt = document.getElementById('ai-exit-prompt');

        // --- Mapeamento de Unidades ---
        const unitMap = {
            'pc': 'Peça', 'pç': 'Peça',
            'm': 'Metros', 'mt': 'Metros',
            'l': 'Litro', 'lt': 'Litro',
            'kg': 'Quilo',
            'un': 'Unidade', 'und': 'Unidade',
            'cx': 'Caixa',
            'pct': 'Pacote',
            'rl': 'Rolo',
            'gl': 'Galão',
            'sc': 'Saco'
        };

        const normalizeUnit = (rawUnit) => {
            const unit = (rawUnit || '').toLowerCase().trim();
            return unitMap[unit] || rawUnit || 'Unidade';
        };

        const VALID_PRODUCT_GROUPS = [
            'Elétrico',
            'Hidráulico',
            'Consumível',
            'Ferramentas Manuais',
            'Material de Corte e Solda',
            'Escritório',
            'Segurança',
            'Outros'
        ];

        const inferProductGroup = (product) => {
            const text = `${product?.name || ''} ${product?.codeRM || ''} ${product?.location || ''}`.toLowerCase();

            if (/cabo|f[íi]o|interruptor|tomada|disjuntor|l[âa]mpada|contator|rel[eé]|el[eé]tr/i.test(text)) {
                return 'Elétrico';
            }
            if (/tubo|conex[aã]o|registro|hidraul|v[áa]lvula|torneira|mangueira|joelho|luva/i.test(text)) {
                return 'Hidráulico';
            }
            if (/luva|epi|capacete|bota|[óo]culos|protetor|máscara|mascara|seguran/i.test(text)) {
                return 'Segurança';
            }
            if (/solda|eletrodo|ma[çc]arico|oxicorte|disco\s*de\s*corte|arame\s*de\s*solda|estanho|fluxo\s*de\s*solda/i.test(text)) {
                return 'Material de Corte e Solda';
            }
            if (/chave|alicate|martelo|torqu[eê]s|trena|estilete|chave\s*de\s*fenda|chave\s*philips|ferramenta\s*manual/i.test(text)) {
                return 'Ferramentas Manuais';
            }
            if (/broca|furadeira|serra\s*(copo|tico|sabre)?|esmerilhadeira|lixadeira/i.test(text)) {
                return 'Ferramentas Manuais';
            }
            if (/papel|caneta|grampo|toner|escrit/i.test(text)) {
                return 'Escritório';
            }
            if (/cola|fita|parafuso|porca|arruela|consum/i.test(text)) {
                return 'Consumível';
            }

            return 'Outros';
        };

        const isMissingGroup = (group) => {
            const normalized = (group || '').toString().trim().toLowerCase();
            return !normalized || normalized === 'n/a' || normalized === '#n/d' || normalized === 'nd';
        };

        const needsGroupAutoFix = (group) => {
            const normalized = (group || '').toString().trim().toLowerCase();
            return isMissingGroup(group) || normalized === 'outros' || normalized === 'ferramentas';
        };

        const autoFillMissingProductGroups = async () => {
            if (hasAutoFilledMissingGroups || isAutoFillingMissingGroups) return;
            if (!hasPermission('update')) return;

            const candidates = products.filter(p => needsGroupAutoFix(p.group));
            if (!candidates.length) {
                hasAutoFilledMissingGroups = true;
                return;
            }

            isAutoFillingMissingGroups = true;

            try {
                const batch = writeBatch(db);

                let changedCount = 0;

                candidates.forEach((product) => {
                    const productRef = doc(productsCollectionRef, product.id);
                    const currentGroup = (product.group || '').toString().trim();
                    const currentGroupNormalized = currentGroup.toLowerCase();
                    const inferredGroup = currentGroupNormalized === 'ferramentas'
                        ? 'Ferramentas Manuais'
                        : inferProductGroup(product);
                    const safeGroup = VALID_PRODUCT_GROUPS.includes(inferredGroup) ? inferredGroup : 'Outros';

                    if (currentGroup === safeGroup) return;

                    batch.update(productRef, { group: safeGroup });
                    changedCount++;
                });

                if (!changedCount) {
                    hasAutoFilledMissingGroups = true;
                    isAutoFillingMissingGroups = false;
                    return;
                }

                await batch.commit();
                showToast(`✅ ${changedCount} item(ns) sem grupo/"Outros"/"Ferramentas" foram atualizados.`, false);
            } catch (error) {
                console.error('Erro ao preencher grupos faltantes:', error);
                showToast('Não foi possível preencher os grupos automaticamente.', true);
            } finally {
                isAutoFillingMissingGroups = false;
                hasAutoFilledMissingGroups = true;
            }
        };

        // --- Gemini API ---
        const callGeminiAPI = async (prompt, model, config = {}) => {
            const apiKey = "";
            if (!apiKey) {
                showToast('Recurso de IA não disponível. Configure uma API Key do Gemini nas configurações.', false);
                return null;
            }
            const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
            const payload = { contents: [{ parts: [{ text: prompt }] }], ...config };
            let retries = 3;
            let delay = 1000;
            while (retries > 0) {
                try {
                    const response = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
                    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                    const result = await response.json();
                    if (result.candidates?.[0]?.content?.parts?.[0]) {
                        return result.candidates[0].content.parts[0].text;
                    } else {
                        console.error('Invalid response structure from Gemini API:', result);
                        return null;
                    }
                } catch (error) {
                    console.error(`API call failed: ${error.message}. Retrying in ${delay}ms...`);
                    retries--;
                    if (retries === 0) {
                        showToast("Erro ao comunicar com o assistente de IA.", true);
                        return null;
                    }
                    await new Promise(res => setTimeout(res, delay));
                    delay *= 2;
                }
            }
        };

        const callTTSAPI = async (text) => {
            const apiKey = "";
            if (!apiKey) return null;
            const apiUrl = `https://texttospeech.googleapis.com/v1/text:synthesize?key=${apiKey}`;
            const payload = {
                input: { text: text },
                voice: { languageCode: 'pt-BR', name: 'pt-BR-Wavenet-B' }, // Voz feminina
                audioConfig: { audioEncoding: 'LINEAR16', sampleRateHertz: 24000 }
            };
            try {
                const response = await fetch(apiUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const result = await response.json();
                return result.audioContent;
            } catch (error) {
                console.error("TTS generation failed:", error);
            }
            return null;
        };

        // --- Funções Auxiliares de Áudio ---
        function base64ToArrayBuffer(base64) {
            const binaryString = window.atob(base64);
            const len = binaryString.length;
            const bytes = new Uint8Array(len);
            for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            return bytes.buffer;
        }

        function pcmToWav(pcmData, sampleRate) {
            const numChannels = 1;
            const bitsPerSample = 16;
            const byteRate = sampleRate * numChannels * (bitsPerSample / 8);
            const blockAlign = numChannels * (bitsPerSample / 8);
            const dataSize = pcmData.length * 2;
            const buffer = new ArrayBuffer(44 + dataSize);
            const view = new DataView(buffer);

            view.setUint32(0, 0x52494646, false); // "RIFF"
            view.setUint32(4, 36 + dataSize, true);
            view.setUint32(8, 0x57415645, false); // "WAVE"
            view.setUint32(12, 0x666d7420, false); // "fmt "
            view.setUint32(16, 16, true);
            view.setUint16(20, 1, true);
            view.setUint16(22, numChannels, true);
            view.setUint32(24, sampleRate, true);
            view.setUint32(28, byteRate, true);
            view.setUint16(32, blockAlign, true);
            view.setUint16(34, bitsPerSample, true);
            view.setUint32(36, 0x64617461, false); // "data"
            view.setUint32(40, dataSize, true);

            for (let i = 0; i < pcmData.length; i++) {
                view.setInt16(44 + i * 2, pcmData[i], true);
            }

            return new Blob([view], { type: 'audio/wav' });
        }


        // --- Funções de UI ---
        const showLoader = (show) => {
            loaderOverlay.style.opacity = show ? '1' : '0';
            loaderOverlay.style.pointerEvents = show ? 'auto' : 'none';
        };

        // 🎨 Skeleton loader para lista de produtos
        const showProductsSkeleton = (show = true, count = 5) => {
            if (!show) {
                productList.innerHTML = '';
                return;
            }

            productList.innerHTML = '';
            for (let i = 0; i < count; i++) {
                const tr = document.createElement('tr');
                tr.className = 'border-b border-slate-200';
                tr.innerHTML = `
                    <td class="p-4"><div class="skeleton w-4 h-4"></div></td>
                    <td class="p-4">
                        <div class="skeleton skeleton-title"></div>
                        <div class="skeleton skeleton-text w-3/4"></div>
                        <div class="skeleton skeleton-text w-1/2"></div>
                    </td>
                    <td class="p-4"><div class="skeleton skeleton-text w-16"></div></td>
                    <td class="p-4"><div class="skeleton skeleton-text w-12"></div></td>
                    <td class="p-4"><div class="skeleton skeleton-text w-20"></div></td>
                    <td class="p-4"><div class="skeleton skeleton-text w-12"></div></td>
                    <td class="p-4"><div class="skeleton skeleton-text w-32"></div></td>
                    <td class="p-4"><div class="skeleton w-20 h-8"></div></td>
                    <td class="p-4"><div class="skeleton w-24 h-8"></div></td>
                `;
                productList.appendChild(tr);
            }
        };

        // 🔐 Funções de Autenticação e Permissões
        const PERMISSIONS = {
            admin: ['create', 'read', 'update', 'delete', 'export', 'import', 'manage_users', 'settings'],
            operador: ['create', 'read', 'update', 'export'],
            visualizador: ['read', 'export']
        };

        const hasPermission = (action) => {
            if (!userRole) return false;
            return PERMISSIONS[userRole]?.includes(action) || false;
        };

        const getUserRole = async (userId) => {
            try {
                const userDocRef = doc(usersCollectionRef, userId);
                const userDoc = await getDoc(userDocRef);
                
                if (userDoc.exists()) {
                    return userDoc.data().role || 'visualizador';
                } else {
                    // Criar usuário novo como visualizador por padrão
                    await setDoc(userDocRef, {
                        email: currentUser.email,
                        displayName: currentUser.displayName,
                        photoURL: currentUser.photoURL,
                        role: 'visualizador',
                        createdAt: serverTimestamp()
                    });
                    showToast('⚠️ Sua conta foi criada como Visualizador. Contate um administrador para alterar permissões.', false);
                    return 'visualizador';
                }
            } catch (error) {
                console.error('Erro ao buscar role do usuário:', error);
                return 'visualizador';
            }
        };

        const updateUIBasedOnPermissions = () => {
            const createButtons = document.querySelectorAll('#add-product-btn, #create-requisition-btn, #add-location-btn');
            const editButtons = document.querySelectorAll('.edit-btn, .delete-btn, .adjust-btn');
            const importExportButtons = document.querySelectorAll('#import-btn, #export-btn, #backup-btn, #restore-btn');
            
            // Ocultar botões baseado em permissões
            createButtons.forEach(btn => {
                btn.style.display = hasPermission('create') ? '' : 'none';
            });

            if (!hasPermission('update')) {
                document.addEventListener('click', (e) => {
                    if (e.target.closest('.edit-btn, .delete-btn, .adjust-btn')) {
                        e.stopPropagation();
                        showToast('🔒 Você não tem permissão para editar/excluir', true);
                        return false;
                    }
                }, true);
            }

            importExportButtons.forEach(btn => {
                if (btn.id === 'export-btn' && hasPermission('export')) {
                    btn.style.display = '';
                } else if ((btn.id === 'import-btn' || btn.id === 'backup-btn' || btn.id === 'restore-btn') && hasPermission('import')) {
                    btn.style.display = '';
                } else if (!hasPermission('export') && !hasPermission('import')) {
                    btn.style.display = 'none';
                }
            });

            // Configurações apenas para admin
            if (settingsBtn) {
                settingsBtn.style.display = hasPermission('settings') ? '' : 'none';
            }
        };

        const showToast = (message, isError = false) => {
            const toast = document.getElementById('toast');
            toast.textContent = message;
            toast.className = isError 
                ? 'bg-red-600 text-white' 
                : 'bg-slate-900 text-white';
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 3000);
        };

        const openModal = (modalId) => {
            document.getElementById(`${modalId}-backdrop`).classList.add('show');
            document.getElementById(modalId).classList.add('show');
        };
        const closeModal = (modalId) => {
            if (currentAudio) {
                currentAudio.pause();
                currentAudio = null;
            }
            const modal = document.getElementById(modalId);
            if (!modal) return;
            const backdrop = document.getElementById(`${modalId}-backdrop`);
            if (backdrop) backdrop.classList.remove('show');
            modal.classList.remove('show');
            modal.classList.remove('max-w-4xl');
            modal.classList.add('max-w-md');
        };

        const showConfirmationModal = (title, message, onConfirm) => {
            const content = `
                <h2 class="text-2xl font-bold mb-4">${title}</h2>
                <p class="text-slate-600 mb-8">${message}</p>
                <div class="flex justify-end gap-4">
                    <button id="confirm-cancel" class="px-6 py-2 bg-slate-200 text-slate-800 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                    <button id="confirm-ok" class="px-6 py-2 bg-red-600 text-white rounded-lg font-semibold hover:bg-red-700 transition">Confirmar</button>
                </div>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            openModal('generic-modal');
            document.getElementById('confirm-ok').onclick = () => { onConfirm(); closeModal('generic-modal'); };
            document.getElementById('confirm-cancel').onclick = () => closeModal('generic-modal');
        };

        const updateAppSettingsUI = (settings) => {
            appSettings = { ...appSettings, ...settings };
            document.getElementById('app-title').textContent = appSettings.appName;
            document.title = appSettings.appName;
            const sidebarTitle = document.getElementById('sidebar-app-title');
            if (sidebarTitle) sidebarTitle.textContent = appSettings.appName;
            const logoContainer = document.getElementById('app-logo');
            if (appSettings.logoUrl) {
                logoContainer.innerHTML = `<img src="${appSettings.logoUrl}" alt="Logo">`;
                logoContainer.className = 'p-0 rounded-lg shadow-lg flex items-center justify-center';
            } else {
                logoContainer.innerHTML = `<svg class="w-8 h-8 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 7v10c0 2.21 3.582 4 8 4s8-1.79 8-4V7M4 7c0-2.21 3.582-4 8-4s8 1.79 8 4M4 7l8 4.5 8-4.5M12 12l8 4.5"></path></svg>`;
                logoContainer.className = 'bg-gradient-to-br from-indigo-600 to-indigo-700 p-3 rounded-xl shadow-lg shadow-indigo-500/20 flex items-center justify-center h-12 w-12';
            }
        };

        // --- Utilitários de Autenticação ---
        const hashPassword = async (password) => {
            const encoder = new TextEncoder();
            const data = encoder.encode(password);
            const hashBuffer = await crypto.subtle.digest('SHA-256', data);
            return Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
        };

        const normalizeUserId = (username) =>
            username.toLowerCase().trim().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');

        const initializeAppSession = (customUser) => {
            currentUser = customUser;
            userRole = customUser.role;
            if (loginBtn) loginBtn.disabled = false;

            // Configurar referências do Firestore
            productsCollectionRef = collection(db, `/artifacts/${appId}/public/data/products`);
            historyCollectionRef = collection(db, `/artifacts/${appId}/public/data/history`);
            requisitionsCollectionRef = collection(db, `/artifacts/${appId}/public/data/requisitions`);
            locationsCollectionRef = collection(db, `/artifacts/${appId}/public/data/locations`);
            settingsDocRef = doc(db, `/artifacts/${appId}/public/data/app_settings/main`);
            usersCollectionRef = collection(db, `/artifacts/${appId}/public/data/users`);
            
            // ⭐ UHE Estrela collections
            estrelaProductsRef = collection(db, `/artifacts/${appId}/public/data/estrela_products`);
            estrelaEntriesRef = collection(db, `/artifacts/${appId}/public/data/estrela_entries`);
            estrelaExitsRef = collection(db, `/artifacts/${appId}/public/data/estrela_exits`);

            // Avatar com iniciais
            const initials = customUser.displayName.split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase();
            const roleLabels = {
                admin: { text: 'Administrador', color: '#6600cc' },
                operador: { text: 'Operador', color: '#005bbf' },
                visualizador: { text: 'Visualizador', color: '#414754' }
            };
            const roleInfo = roleLabels[userRole] || roleLabels.operador;

            if (userInfoDiv) userInfoDiv.innerHTML = `
                <div class="flex items-center gap-2">
                    <div class="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-bold flex-shrink-0" style="background:#005bbf;">${initials}</div>
                    <div class="text-left hidden md:block">
                        <p class="text-sm font-semibold" style="color:#191c1d;">${customUser.displayName}</p>
                        <p class="text-xs font-bold px-2 py-0.5 rounded-full inline-block" style="background:rgba(0,91,191,0.08); color:${roleInfo.color};">${roleInfo.text}</p>
                    </div>
                </div>`;

            // Sidebar footer
            const sidebarName = document.getElementById('sidebar-user-name');
            const sidebarRole = document.getElementById('sidebar-user-role');
            const sidebarAvatar = document.getElementById('sidebar-avatar');
            if (sidebarName) sidebarName.textContent = customUser.displayName.split(' ')[0];
            if (sidebarRole) sidebarRole.textContent = roleInfo.text;
            if (sidebarAvatar) {
                sidebarAvatar.innerHTML = `<div class="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-bold" style="background:#005bbf;">${initials}</div>`;
                sidebarAvatar.className = 'flex-shrink-0';
            }

            loginScreen.classList.add('hidden');
            appContainer.classList.remove('hidden');
            if (logoutBtn) logoutBtn.classList.remove('hidden');

            updateUIBasedOnPermissions();
            setupListeners();
            isAuthInitialized = true;
        };

        // --- Lógica de Autenticação e Inicialização (sem Firebase Auth) ---
        const savedSession = localStorage.getItem('appUser');
        if (savedSession) {
            try {
                const customUser = JSON.parse(savedSession);
                initializeAppSession(customUser);
            } catch(e) {
                localStorage.removeItem('appUser');
            }
        }

        const handleFirestoreError = (error, dataType) => {
            console.error(`Erro ao buscar ${dataType}:`, error);

            let errorMessage = '';
            switch (error.code) {
                case 'permission-denied':
                    errorMessage = `🔒 Sem permissão para acessar ${dataType}. Configure as regras do Firebase.`;
                    authStatusDiv.innerHTML = `<p class="text-sm text-red-600 font-semibold">Sem Permissão</p><p class="text-xs text-slate-400">Verifique Firebase</p>`;
                    break;
                case 'unavailable':
                    errorMessage = `📡 Servidor indisponível. Verifique sua conexão.`;
                    break;
                case 'not-found':
                    errorMessage = `❓ Coleção de ${dataType} não encontrada.`;
                    break;
                default:
                    errorMessage = `❌ Erro ao carregar ${dataType}. Tente recarregar a página.`;
            }

            showToast(errorMessage, true);
        };

        const stopSnapshotGroup = (unsubscribers) => {
            unsubscribers.forEach(unsub => {
                try { if (typeof unsub === 'function') unsub(); } catch (_) {}
            });
            unsubscribers.length = 0;
        };

        function stopCoreListeners() {
            stopSnapshotGroup(coreUnsubscribers);
        }

        function stopEstrelaListeners() {
            stopSnapshotGroup(estrelaUnsubscribers);
        }

        function startCoreListeners() {
            if (coreUnsubscribers.length > 0) return;

            coreUnsubscribers.push(onSnapshot(settingsDocRef, (doc) => {
                if (doc.exists()) {
                    const data = doc.data();
                    if (data.appName === 'Estoque Taboca' || data.appName === 'Estoque Estrela') {
                        data.appName = 'UHE Estrela';
                        setDoc(settingsDocRef, { appName: 'UHE Estrela' }, { merge: true }).catch(() => {});
                    }
                    updateAppSettingsUI(data);
                }
                else updateAppSettingsUI({ appName: 'UHE Estrela', logoUrl: null });
            }, (error) => handleFirestoreError(error, 'configurações')));

            coreUnsubscribers.push(onSnapshot(productsCollectionRef, (snapshot) => {
                products = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

                if (!hasAutoFilledMissingGroups) {
                    autoFillMissingProductGroups();
                }

                if (isDataLoaded) {
                    renderProducts();
                    updateDashboard();
                    renderEntryView();
                    renderExitView();
                }
            }, (error) => handleFirestoreError(error, 'produtos')));

            coreUnsubscribers.push(onSnapshot(historyCollectionRef, (snapshot) => {
                history = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (isDataLoaded) {
                    updateDashboard();
                    renderExitLog(exitsSearchInput.value);
                    renderRMView(rmSearchInput.value);
                }
            }, (error) => handleFirestoreError(error, 'histórico')));

            coreUnsubscribers.push(onSnapshot(requisitionsCollectionRef, (snapshot) => {
                requisitions = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (isDataLoaded) renderRequisitions();
            }, (error) => handleFirestoreError(error, 'requisições')));

            coreUnsubscribers.push(onSnapshot(locationsCollectionRef, (snapshot) => {
                locations = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (!isDataLoaded) {
                    isDataLoaded = true;
                    showLoader(false);
                    showProductsSkeleton(false); // 🎨 Remover skeleton ao carregar
                    renderProducts();
                    updateDashboard();
                    renderRequisitions();
                    renderLocations();
                    renderRMView();
                } else {
                    renderLocations();
                }
            }, (error) => handleFirestoreError(error, 'locais')));
        }

        function startEstrelaListeners() {
            if (estrelaUnsubscribers.length > 0) return;

            estrelaUnsubscribers.push(onSnapshot(estrelaProductsRef, (snapshot) => {
                estrelaProducts = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
                if (isDataLoaded && currentViewId === 'estrela-view') renderEstrelaEstoque();
            }, (error) => handleFirestoreError(error, 'produtos Estrela')));

            estrelaUnsubscribers.push(onSnapshot(estrelaEntriesRef, (snapshot) => {
                estrelaEntries = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
                if (isDataLoaded && currentViewId === 'estrela-view') renderEstrelaEntradas();
            }, (error) => handleFirestoreError(error, 'entradas Estrela')));

            estrelaUnsubscribers.push(onSnapshot(estrelaExitsRef, (snapshot) => {
                estrelaExits = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
                if (isDataLoaded && currentViewId === 'estrela-view') renderEstrelaSaidas();
            }, (error) => handleFirestoreError(error, 'saídas Estrela')));
        }

        function syncRealtimeListeners() {
            if (!currentUser || document.hidden) {
                stopEstrelaListeners();
                stopCoreListeners();
                return;
            }

            startCoreListeners();
            if (currentViewId === 'estrela-view') startEstrelaListeners();
            else stopEstrelaListeners();
        }

        function setupListeners() {
            syncRealtimeListeners();
            if (!hasVisibilityListener) {
                document.addEventListener('visibilitychange', syncRealtimeListeners);
                hasVisibilityListener = true;
            }
        }

        // --- Funções de Lógica (Firestore) ---
        const addHistoryEntry = async (productId, type, quantity, newTotal, details = {}, productData = null) => {
            const product = productData || products.find(p => p.id === productId);

            if (!product && type !== 'Edição') {
                console.warn(`Tentativa de adicionar histórico para produto não encontrado (ID: ${productId})`);
                return;
            }

            try {
                await addDoc(historyCollectionRef, {
                    productId,
                    productCode: product?.code || 'N/A',
                    productCodeRM: product?.codeRM || 'N/A',
                    productName: product?.name || 'Produto Excluído',
                    type, quantity: Math.abs(quantity), newTotal,
                    withdrawnBy: details.withdrawnBy || null,
                    teamLeader: details.teamLeader || null,
                    applicationLocation: details.applicationLocation || null,
                    obra: details.obra || null,
                    details: details.details || null,
                    rmProcessed: false,
                    date: serverTimestamp(),
                    performedBy: currentUser?.displayName || 'Anônimo', // 🔍 Auditoria: usuário que realizou ação
                    timestamp: serverTimestamp() // 🔍 Auditoria: timestamp preciso
                });
            } catch (error) {
                console.error("Erro ao adicionar registro no histórico: ", error);
                showToast("Falha ao salvar histórico.", true);
            }
        };
        
        const updateDashboard = () => {
            const totalProducts = products.length;
            const totalUnits = products.reduce((sum, p) => sum + (p.quantity || 0), 0);
            const lowStockItems = products.filter(p => p.quantity <= p.minQuantity).length;
            const pendingReqs = requisitions.filter(r => r.status === 'Pendente' || r.status === 'pending').length;

            const timestampToMillis = (entry) => {
                if (!entry?.date) return 0;
                if (typeof entry.date.seconds === 'number') return entry.date.seconds * 1000;
                if (typeof entry.date.toMillis === 'function') return entry.date.toMillis();
                return 0;
            };

            const getTopExitedItems = (days, limit = 5) => {
                const cutoff = Date.now() - (days * 24 * 60 * 60 * 1000);
                const map = new Map();

                history
                    .filter(h => (h.type === 'Saída' || h.type === 'Saída por Requisição') && timestampToMillis(h) >= cutoff)
                    .forEach(h => {
                        const key = h.productId || h.productCode || h.productCodeRM || h.productName;
                        const previous = map.get(key) || {
                            productName: h.productName || 'Produto sem nome',
                            productCode: h.productCodeRM || h.productCode || 'N/A',
                            totalQty: 0
                        };
                        previous.totalQty += Math.abs(Number(h.quantity) || 0);
                        map.set(key, previous);
                    });

                return [...map.values()].sort((a, b) => b.totalQty - a.totalQty).slice(0, limit);
            };

            const getDailyExitSeries = (days = 7) => {
                const labels = [];
                const values = [];
                const now = new Date();
                const map = new Map();

                for (let i = days - 1; i >= 0; i--) {
                    const date = new Date(now);
                    date.setHours(0, 0, 0, 0);
                    date.setDate(now.getDate() - i);
                    const key = date.toISOString().slice(0, 10);
                    map.set(key, 0);
                    labels.push(date.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' }));
                }

                history
                    .filter(h => h.type === 'Saída' || h.type === 'Saída por Requisição')
                    .forEach((h) => {
                        const millis = timestampToMillis(h);
                        if (!millis) return;
                        const dt = new Date(millis);
                        dt.setHours(0, 0, 0, 0);
                        const key = dt.toISOString().slice(0, 10);
                        if (!map.has(key)) return;
                        map.set(key, map.get(key) + Math.abs(Number(h.quantity) || 0));
                    });

                map.forEach((val) => values.push(val));
                return { labels, values };
            };

            const top3Days = getTopExitedItems(3);
            const top7Days = getTopExitedItems(7);
            const top30Days = getTopExitedItems(30, 8);
            const dailySeries = getDailyExitSeries(7);

            const renderTopList = (items) => {
                if (items.length === 0) {
                    return `<p class="text-xs" style="color:#727785;">Sem saídas registradas no período.</p>`;
                }

                return `
                    <div class="space-y-2">
                        ${items.map((item, index) => `
                            <div class="flex items-center justify-between gap-3 p-2 rounded-lg" style="background:#f8f9fa;">
                                <div class="min-w-0">
                                    <p class="text-sm font-semibold truncate" style="color:#191c1d;">${index + 1}. ${item.productName}</p>
                                    <p class="text-xs truncate" style="color:#727785;">${item.productCode}</p>
                                </div>
                                <span class="text-sm font-extrabold" style="color:#ba1a1a;">${item.totalQty}</span>
                            </div>
                        `).join('')}
                    </div>
                `;
            };

            // Atualizar badges de requisições (sidebar + bottom nav)
            [document.getElementById('nav-req-badge'), document.getElementById('bottom-req-badge')].forEach(badge => {
                if (!badge) return;
                if (pendingReqs > 0) {
                    badge.textContent = pendingReqs > 99 ? '99+' : pendingReqs;
                    badge.classList.remove('hidden');
                } else {
                    badge.classList.add('hidden');
                }
            });

            // Bento Grid - Sovereign Ledger style
            dashboardStats.innerHTML = `
                <!-- KPI Principal: Total de Unidades -->
                <div class="col-span-2 bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="grid-column: span 2; animation-delay: 0.02s;">
                    <div class="flex justify-between items-start mb-4">
                        <div class="p-2.5 rounded-xl" style="background:rgba(0,91,191,0.1);">
                            <span class="material-symbols-outlined" style="font-size:22px; color:#005bbf; font-variation-settings:'FILL' 0,'wght' 400;">shelves</span>
                        </div>
                        <span class="text-xs font-bold uppercase tracking-wider px-2.5 py-1 rounded-lg" style="background:rgba(0,91,191,0.07); color:#005bbf;">Total Geral</span>
                    </div>
                    <p class="text-4xl font-extrabold tracking-tighter" style="color:#191c1d;">${totalUnits.toLocaleString('pt-BR')}</p>
                    <p class="text-sm font-medium mt-1" style="color:#727785;">Unidades em Estoque</p>
                    <button onclick="downloadEstrelaExcel()" class="mt-4 w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-sm font-bold transition-all hover:opacity-90" style="background:linear-gradient(135deg, #fbbc05 0%, #f9a825 100%); color:#261a00; box-shadow:0 4px 12px rgba(251,188,5,0.3);">
                        <span class="material-symbols-outlined" style="font-size:20px; font-variation-settings:'FILL' 1;">download</span>
                        Baixar Controle Estoque UHE Estrela
                    </button>
                </div>
                <!-- KPI: Produtos (SKUs) -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="animation-delay: 0.08s;">
                    <div class="p-2.5 rounded-xl w-fit mb-3" style="background:rgba(0,110,44,0.1);">
                        <span class="material-symbols-outlined" style="font-size:22px; color:#006e2c; font-variation-settings:'FILL' 0,'wght' 400;">category</span>
                    </div>
                    <p class="text-3xl font-bold tracking-tight" style="color:#191c1d;">${totalProducts}</p>
                    <p class="text-xs font-semibold tracking-tight mt-0.5" style="color:#727785;">Produtos (SKUs)</p>
                </div>
                <!-- KPI: Estoque Baixo -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="border-bottom: 3px solid #fbbc05; animation-delay: 0.14s;">
                    <div class="p-2.5 rounded-xl w-fit mb-3" style="background:rgba(121,89,0,0.1);">
                        <span class="material-symbols-outlined" style="font-size:22px; color:#795900; font-variation-settings:'FILL' 0,'wght' 400;">warning</span>
                    </div>
                    <p class="text-3xl font-bold tracking-tight" style="color:#191c1d;">${lowStockItems}</p>
                    <p class="text-xs font-semibold tracking-tight mt-0.5" style="color:#727785;">Itens com Estoque Baixo</p>
                </div>
                <!-- KPI: Requisições Pendentes -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="animation-delay: 0.2s;">
                    <div class="p-2.5 rounded-xl w-fit mb-3" style="background:rgba(0,91,191,0.08);">
                        <span class="material-symbols-outlined" style="font-size:22px; color:#005bbf; font-variation-settings:'FILL' 0,'wght' 400;">assignment_late</span>
                    </div>
                    <p class="text-3xl font-bold tracking-tight" style="color:#191c1d;">${pendingReqs}</p>
                    <p class="text-xs font-semibold tracking-tight mt-0.5" style="color:#727785;">Req. Pendentes</p>
                </div>
                <!-- Ranking: Mais Saíram (3 dias) -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="animation-delay: 0.26s;">
                    <div class="flex items-center justify-between mb-3">
                        <h4 class="text-sm font-bold uppercase tracking-wide" style="color:#191c1d;">Top Saídas (3 dias)</h4>
                        <span class="material-symbols-outlined" style="font-size:20px; color:#ba1a1a;">trending_down</span>
                    </div>
                    ${renderTopList(top3Days)}
                </div>
                <!-- Ranking: Mais Saíram (7 dias) -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="animation-delay: 0.32s;">
                    <div class="flex items-center justify-between mb-3">
                        <h4 class="text-sm font-bold uppercase tracking-wide" style="color:#191c1d;">Top Saídas (7 dias)</h4>
                        <span class="material-symbols-outlined" style="font-size:20px; color:#795900;">calendar_month</span>
                    </div>
                    ${renderTopList(top7Days)}
                </div>`;

            // Ajustar grid do container de stats para bento
            dashboardStats.className = 'grid grid-cols-2 gap-3 md:gap-4';
            const recentHistory = [...history].sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0)).slice(0, 5);
            let activityHTML = `
                <div class="grid grid-cols-1 xl:grid-cols-3 gap-4 mt-6 mb-5">
                    <div class="dashboard-chart-card dashboard-reveal xl:col-span-2" style="animation-delay: 0.05s;">
                        <div class="flex items-center justify-between mb-3">
                            <h3 class="text-sm font-bold uppercase tracking-wide" style="color:#191c1d;">Giro Diário (7 dias)</h3>
                            <span class="dashboard-top-pill">Saídas</span>
                        </div>
                        <div class="h-64 md:h-72">
                            <canvas id="dashboard-trend-chart"></canvas>
                        </div>
                    </div>
                    <div class="dashboard-chart-card dashboard-reveal" style="animation-delay: 0.1s;">
                        <div class="flex items-center justify-between mb-3">
                            <h3 class="text-sm font-bold uppercase tracking-wide" style="color:#191c1d;">Mais Rodados (30 dias)</h3>
                            <span class="dashboard-top-pill">Top ${top30Days.length}</span>
                        </div>
                        <div class="h-64 md:h-72">
                            <canvas id="dashboard-top-items-chart"></canvas>
                        </div>
                    </div>
                </div>
                <div class="bg-white border border-slate-200 rounded-2xl p-4 mb-5 dashboard-reveal" style="animation-delay: 0.14s; box-shadow: 0 10px 30px rgba(25, 28, 29, 0.04);">
                    <div class="flex items-center justify-between mb-3">
                        <h3 class="text-sm font-bold uppercase tracking-wide" style="color:#191c1d;">Itens com Maior Giro</h3>
                        <button onclick="document.querySelector('[data-view=exit-log-view]').click()" class="text-xs font-bold uppercase tracking-widest" style="color:#005bbf; letter-spacing:0.08em;">Detalhar</button>
                    </div>
                    <div class="space-y-2">
                        ${top30Days.length
                            ? top30Days.slice(0, 5).map((item, index) => `
                                <div class="flex items-center justify-between p-2.5 rounded-lg" style="background:#f8f9fa;">
                                    <div class="min-w-0">
                                        <p class="text-sm font-semibold truncate" style="color:#191c1d;">${index + 1}. ${item.productName}</p>
                                        <p class="text-xs truncate" style="color:#727785;">${item.productCode}</p>
                                    </div>
                                    <span class="text-sm font-extrabold" style="color:#ba1a1a;">${item.totalQty}</span>
                                </div>
                            `).join('')
                            : '<p class="text-sm" style="color:#727785;">Sem saídas registradas nos últimos 30 dias.</p>'}
                    </div>
                </div>
                <div class="flex justify-between items-center px-1 mt-6 mb-3">
                    <h3 class="text-lg font-bold tracking-tight" style="color:#191c1d;">Atividade Recente</h3>
                    <button onclick="document.querySelector('[data-view=exit-log-view]').click()" class="text-xs font-bold uppercase tracking-widest" style="color:#005bbf; letter-spacing:0.08em;">Ver tudo</button>
                </div>
                <div class="rounded-2xl p-2 space-y-2" style="background:#f3f4f5;">`;
            if (recentHistory.length > 0) {
                 recentHistory.forEach(h => {
                     let iconBg = '', iconColor = '', iconName = '', quantDisplay = '';
                     switch(h.type) {
                         case 'Ajuste Entrada': case 'Entrada':
                             iconBg = 'rgba(0,110,44,0.12)'; iconColor = '#006e2c'; iconName = 'arrow_upward';
                             quantDisplay = `<p class="text-sm font-black" style="color:#006e2c;">+${h.quantity}</p>`; break;
                         case 'Ajuste Saída':
                             iconBg = 'rgba(121,89,0,0.12)'; iconColor = '#795900'; iconName = 'arrow_downward';
                             quantDisplay = `<p class="text-sm font-black" style="color:#795900;">-${h.quantity}</p>`; break;
                         case 'Saída': case 'Saída por Requisição':
                             iconBg = 'rgba(186,26,26,0.1)'; iconColor = '#ba1a1a'; iconName = 'arrow_downward';
                             quantDisplay = `<p class="text-sm font-black" style="color:#ba1a1a;">-${h.quantity}</p>`; break;
                         case 'Criação': case 'Importação':
                             iconBg = 'rgba(0,91,191,0.1)'; iconColor = '#005bbf'; iconName = 'add_circle';
                             quantDisplay = `<p class="text-sm font-black" style="color:#005bbf;">+${h.quantity}</p>`; break;
                         default:
                             iconBg = 'rgba(65,71,84,0.1)'; iconColor = '#414754'; iconName = 'edit';
                             quantDisplay = `<p class="text-sm font-black" style="color:#414754;">&mdash;</p>`; break;
                     }
                     const dateStr = h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR', {day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'}) : '';
                     activityHTML += `
                         <div class="flex items-center gap-3 p-3 rounded-xl transition-colors" style="background:#fff;">
                             <div class="w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0" style="background:${iconBg};">
                                 <span class="material-symbols-outlined" style="font-size:20px; color:${iconColor};">${iconName}</span>
                             </div>
                             <div class="flex-1 min-w-0">
                                 <p class="text-sm font-bold truncate" style="color:#191c1d;">${h.productName}</p>
                                 <p class="text-xs truncate" style="color:#727785;">${h.nfNumber ? `NF ${h.nfNumber}` : h.type}${h.withdrawnBy ? ' · ' + h.withdrawnBy : ''} · ${dateStr}</p>
                             </div>
                             ${quantDisplay}
                         </div>`;
                 });
            } else {
                activityHTML += `<p class="text-center py-8 text-sm" style="color:#727785;">Nenhuma atividade registrada ainda.</p>`;
            }
            activityHTML += `</div>`;
            dashboardActivity.innerHTML = activityHTML;

            if (typeof Chart !== 'undefined') {
                if (dashboardTopItemsChart) {
                    dashboardTopItemsChart.destroy();
                    dashboardTopItemsChart = null;
                }
                if (dashboardTrendChart) {
                    dashboardTrendChart.destroy();
                    dashboardTrendChart = null;
                }

                const trendCtx = document.getElementById('dashboard-trend-chart');
                if (trendCtx) {
                    dashboardTrendChart = new Chart(trendCtx, {
                        type: 'line',
                        data: {
                            labels: dailySeries.labels,
                            datasets: [{
                                label: 'Saídas',
                                data: dailySeries.values,
                                fill: true,
                                tension: 0.35,
                                borderColor: '#005bbf',
                                backgroundColor: 'rgba(0, 91, 191, 0.12)',
                                pointBackgroundColor: '#005bbf',
                                pointRadius: 3
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: { display: false }
                            },
                            scales: {
                                x: {
                                    grid: { display: false },
                                    ticks: { color: '#727785' }
                                },
                                y: {
                                    beginAtZero: true,
                                    grid: { color: 'rgba(114, 119, 133, 0.12)' },
                                    ticks: { color: '#727785' }
                                }
                            }
                        }
                    });
                }

                const topItemsCtx = document.getElementById('dashboard-top-items-chart');
                if (topItemsCtx) {
                    dashboardTopItemsChart = new Chart(topItemsCtx, {
                        type: 'bar',
                        data: {
                            labels: top30Days.slice(0, 6).map(item => item.productName.length > 18 ? `${item.productName.slice(0, 18)}...` : item.productName),
                            datasets: [{
                                label: 'Qtd',
                                data: top30Days.slice(0, 6).map(item => item.totalQty),
                                backgroundColor: ['#005bbf', '#1a73e8', '#006e2c', '#fbbc05', '#e65100', '#795900'],
                                borderRadius: 8,
                                maxBarThickness: 24
                            }]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            indexAxis: 'y',
                            plugins: {
                                legend: { display: false }
                            },
                            scales: {
                                x: {
                                    beginAtZero: true,
                                    grid: { color: 'rgba(114, 119, 133, 0.12)' },
                                    ticks: { color: '#727785' }
                                },
                                y: {
                                    grid: { display: false },
                                    ticks: { color: '#414754' }
                                }
                            }
                        }
                    });
                }
            }
        };

        const normalizeSearchText = (value = '') => {
            return String(value)
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase()
                .trim();
        };

        const buildProductSearchText = (p) => {
            return [p.name, p.codeRM, p.code, p.location, p.group, p.unit].filter(Boolean).join(' ');
        };

        const generateNextProductCode = () => {
            if (products.length === 0) return '1';
            const maxCode = products.reduce((max, p) => {
                const codeNum = parseInt(p.code, 10);
                return !isNaN(codeNum) && codeNum > max ? codeNum : max;
            }, 0);
            return (maxCode + 1).toString();
        };

        const generateNextRequisitionNumber = () => {
            if (requisitions.length === 0) return '00001';
            const maxNum = requisitions.reduce((max, r) => {
                const reqNum = parseInt(r.number, 10);
                return !isNaN(reqNum) && reqNum > max ? reqNum : max;
            }, 0);
            return (maxNum + 1).toString().padStart(5, '0');
        };

        // --- Funções de Renderização ---
        const renderProducts = () => {
            const filterText = searchInput.value.toLowerCase();
            let processedProducts = [...products]; 
            if (inventoryFilter === 'low_stock') {
                processedProducts = processedProducts.filter(p => p.quantity <= p.minQuantity);
            }
            if (filterText) {
                processedProducts = processedProducts.filter(p =>
                    (p.name && p.name.toLowerCase().includes(filterText)) || 
                    (p.code && p.code.toLowerCase().includes(filterText)) ||
                    (p.codeRM && p.codeRM.toLowerCase().includes(filterText))
                );
            }
            if (inventorySortOrder === 'name_asc') processedProducts.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            else if (inventorySortOrder === 'code_asc') processedProducts.sort((a, b) => (a.codeRM || '').localeCompare(b.codeRM || ''));
            else if (inventorySortOrder === 'location_asc') processedProducts.sort((a, b) => (a.location || '').localeCompare(b.location || ''));

            // 🚀 Paginação
            const totalProducts = processedProducts.length;
            const totalPages = Math.ceil(totalProducts / itemsPerPage);
            const startIndex = (currentPage - 1) * itemsPerPage;
            const endIndex = startIndex + itemsPerPage;
            const paginatedProducts = processedProducts.slice(startIndex, endIndex);

            productList.innerHTML = '';
            noProductsMessage.classList.toggle('hidden', !(processedProducts.length === 0 && isDataLoaded));
            if (processedProducts.length === 0 && isDataLoaded) {
                noProductsMessage.innerHTML = `<p class="mb-4 text-lg">Nenhum produto encontrado.</p><button id="go-to-add-product" class="bg-indigo-600 text-white font-semibold px-5 py-2 rounded-lg hover:bg-indigo-700 transition">Cadastrar primeiro produto</button>`;
            }

            // Renderizar apenas produtos da página atual
            paginatedProducts.forEach(p => {
                const isLowStock = p.quantity <= p.minQuantity;
                const isChecked = selectedProductIds.has(p.id);
                const tr = document.createElement('tr');
                tr.className = `hover:bg-slate-50 transition-colors duration-150`;
                tr.innerHTML = `
                    <td class="p-4 text-center"><input type="checkbox" class="product-checkbox h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" data-id="${p.id}" ${isChecked ? 'checked' : ''}></td>
                    <td class="p-4 align-top">
                        <p class="font-bold text-slate-800">${p.name}</p>
                        <p class="text-sm text-slate-500">Cód. RM: <span class="font-medium">${p.codeRM || 'N/A'}</span></p>
                        <p class="text-xs text-slate-400">SKU: ${p.code}</p>
                    </td>
                    <td class="p-4 align-top text-slate-600">${p.group || 'N/A'}</td>
                    <td class="p-4 align-top text-slate-600">${p.unit || 'N/A'}</td>
                    <td class="p-4 align-top"><div class="flex items-center gap-2"><p class="text-lg font-bold text-slate-800">${p.quantity}</p>${isLowStock ? '<span class="px-2 py-0.5 text-xs font-semibold text-yellow-800 bg-yellow-100 rounded-full ml-1">Baixo</span>' : ''}</div></td>
                    <td class="p-4 align-top text-slate-600">${p.minQuantity}</td>
                    <td class="p-4 align-top text-slate-600">${p.location}</td>
                    <td class="p-4 align-top text-center"><div class="flex justify-center items-center gap-1"><button data-id="${p.id}" class="history-btn text-slate-500 hover:text-purple-600 p-2 rounded-full hover:bg-purple-100 transition" title="Histórico"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg></button><button data-id="${p.id}" class="edit-btn text-slate-500 hover:text-blue-600 p-2 rounded-full hover:bg-blue-100 transition" title="Editar"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg></button><button data-id="${p.id}" class="delete-btn text-slate-500 hover:text-red-600 p-2 rounded-full hover:bg-red-100 transition" title="Excluir"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg></button></div></td>`;
                productList.appendChild(tr);
            });

            // 🚀 Renderizar controles de paginação
            renderPagination(totalProducts, totalPages);
            updateSelectionActionButtonsState();
        };

        // 🚀 Função para renderizar paginação
        const renderPagination = (totalProducts, totalPages) => {
            const paginationContainer = document.getElementById('pagination-controls');
            if (!paginationContainer) return;

            if (totalPages <= 1) {
                paginationContainer.innerHTML = `<p class="text-sm text-slate-500">Mostrando ${totalProducts} produtos</p>`;
                return;
            }

            const startItem = (currentPage - 1) * itemsPerPage + 1;
            const endItem = Math.min(currentPage * itemsPerPage, totalProducts);

            paginationContainer.innerHTML = `
                <div class="flex items-center justify-between gap-4">
                    <p class="text-sm text-slate-600">
                        Mostrando <span class="font-semibold">${startItem}-${endItem}</span> de <span class="font-semibold">${totalProducts}</span> produtos
                    </p>
                    <div class="flex gap-2">
                        <button id="prev-page" ${currentPage === 1 ? 'disabled' : ''} 
                            class="px-3 py-1 text-sm font-medium rounded-lg transition ${currentPage === 1 ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-indigo-100 text-indigo-700 hover:bg-indigo-200'}">
                            ← Anterior
                        </button>
                        <span class="px-3 py-1 text-sm font-semibold text-slate-700">
                            Página ${currentPage} de ${totalPages}
                        </span>
                        <button id="next-page" ${currentPage === totalPages ? 'disabled' : ''} 
                            class="px-3 py-1 text-sm font-medium rounded-lg transition ${currentPage === totalPages ? 'bg-slate-100 text-slate-400 cursor-not-allowed' : 'bg-indigo-100 text-indigo-700 hover:bg-indigo-200'}">
                            Próxima →
                        </button>
                    </div>
                </div>
            `;

            // Event listeners para navegação
            document.getElementById('prev-page')?.addEventListener('click', () => {
                if (currentPage > 1) {
                    currentPage--;
                    renderProducts();
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }
            });

            document.getElementById('next-page')?.addEventListener('click', () => {
                if (currentPage < totalPages) {
                    currentPage++;
                    renderProducts();
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }
            });
        };
        
        const renderRequisitions = () => {
            const sortedRequisitions = [...requisitions].sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));
            requisitionsList.innerHTML = '';
            noRequisitionsMessage.classList.toggle('hidden', sortedRequisitions.length > 0);
            if(sortedRequisitions.length === 0 && isDataLoaded) {
                noRequisitionsMessage.innerHTML = `<p class="text-lg">Nenhuma requisição criada ainda.</p>`;
            }

            sortedRequisitions.forEach(req => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                tr.innerHTML = `
                    <td class="p-4 text-center"><input type="checkbox" class="req-checkbox h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" data-id="${req.id}"></td>
                    <td class="p-4 font-bold text-slate-700">${req.number}</td>
                    <td class="p-4 text-slate-600">${req.date ? new Date(req.date.seconds * 1000).toLocaleDateString('pt-BR') : '...'}</td>
                    <td class="p-4 text-slate-600">${req.requester}</td>
                    <td class="p-4 text-slate-600">${req.applicationLocation}</td>
                    <td class="p-4 text-center">
                        <button data-id="${req.id}" class="view-requisition-btn text-indigo-600 hover:underline font-semibold">Detalhes/PDF</button>
                    </td>
                `;
                requisitionsList.appendChild(tr);
            });
        };

        const renderExitLog = (filter = '') => {
            const lowerCaseFilter = filter.toLowerCase();
            const exitHistory = history.filter(h => 
                (h.type === 'Saída' || h.type === 'Saída por Requisição') && 
                ((h.productName && h.productName.toLowerCase().includes(lowerCaseFilter)) || 
                 (h.productCode && h.productCode.toLowerCase().includes(lowerCaseFilter)) ||
                 (h.withdrawnBy && h.withdrawnBy.toLowerCase().includes(lowerCaseFilter)) ||
                 (h.applicationLocation && h.applicationLocation.toLowerCase().includes(lowerCaseFilter)))
            ).sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            exitsList.innerHTML = '';
            noExitsMessage.classList.toggle('hidden', !(exitHistory.length === 0 && isDataLoaded));
            if (exitHistory.length === 0 && isDataLoaded) {
                 noExitsMessage.innerHTML = `<p class="text-lg">Nenhuma saída registrada.</p>`;
            }

            exitHistory.forEach(h => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                tr.innerHTML = `
                    <td class="p-4 text-sm text-slate-600">${h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : '...'}</td>
                    <td class="p-4"><p class="font-semibold text-slate-800">${h.productName}</p><p class="text-sm text-slate-500">${h.productCodeRM || h.productCode}</p></td>
                    <td class="p-4 text-center text-red-600 font-bold text-lg">${h.quantity}</td>
                    <td class="p-4 text-slate-700">${h.withdrawnBy}</td>
                    <td class="p-4 text-slate-700">${h.obra || 'N/A'}</td>
                    <td class="p-4 text-slate-700">${h.applicationLocation}</td>
                `;
                exitsList.appendChild(tr);
            });
        };

        const renderRMView = (filter = '') => {
            const lowerCaseFilter = filter.toLowerCase();
            const exitHistory = history.filter(h => 
                (h.type === 'Saída' || h.type === 'Saída por Requisição') &&
                ((h.productName && h.productName.toLowerCase().includes(lowerCaseFilter)) || 
                 (h.productCodeRM && h.productCodeRM.toLowerCase().includes(lowerCaseFilter)) ||
                 (h.withdrawnBy && h.withdrawnBy.toLowerCase().includes(lowerCaseFilter)) ||
                 (h.applicationLocation && h.applicationLocation.toLowerCase().includes(lowerCaseFilter)))
            ).sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            rmList.innerHTML = '';
            noRmMessage.classList.toggle('hidden', !(exitHistory.length === 0 && isDataLoaded));
            if (exitHistory.length === 0 && isDataLoaded) {
                 noRmMessage.innerHTML = `<p class="text-lg">Nenhuma saída encontrada.</p>`;
            }

            exitHistory.forEach(h => {
                const tr = document.createElement('tr');
                tr.className = `transition-colors duration-200 ${h.rmProcessed ? 'bg-green-50 hover:bg-green-100/70' : 'hover:bg-slate-50'}`;
                
                const locationData = locations.find(loc => loc.name === h.applicationLocation);
                const locationCode = locationData ? locationData.code : '';

                tr.innerHTML = `
                    <td class="p-4 text-sm text-slate-600">${h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : '...'}</td>
                    <td class="p-4"><p class="font-semibold text-slate-800">${h.productName}</p><p class="text-sm text-slate-500">${h.productCodeRM || h.productCode}</p></td>
                    <td class="p-4">
                        <div class="flex items-center gap-1 flex-wrap">
                            <button class="copy-btn text-xs bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold py-1 px-2 rounded" data-copy-text="${h.productName}">Nome</button>
                            <button class="copy-btn text-xs bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold py-1 px-2 rounded" data-copy-text="${h.productCodeRM || ''}">Cód. RM</button>
                            <button class="copy-btn text-xs bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold py-1 px-2 rounded" data-copy-text="${h.applicationLocation || ''}">Local</button>
                            <button class="copy-btn text-xs bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold py-1 px-2 rounded" data-copy-text="${locationCode || ''}">Cód. Local</button>
                        </div>
                    </td>
                    <td class="p-4 text-center text-red-600 font-bold text-lg">${h.quantity}</td>
                    <td class="p-4 text-slate-700">${h.withdrawnBy}</td>
                    <td class="p-4 text-center">
                        <input type="checkbox" class="rm-status-checkbox h-5 w-5 rounded text-indigo-600 focus:ring-indigo-500 cursor-pointer" data-id="${h.id}" ${h.rmProcessed ? 'checked' : ''}>
                    </td>
                `;
                rmList.appendChild(tr);
            });
        };

        const renderLocations = () => {
            const sortedLocations = [...locations].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            locationsList.innerHTML = '';
            noLocationsMessage.classList.toggle('hidden', sortedLocations.length > 0);
            if(sortedLocations.length === 0 && isDataLoaded) {
                noLocationsMessage.innerHTML = `<p class="text-lg">Nenhum local cadastrado.</p>`;
            }

            sortedLocations.forEach(loc => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                const aliasesHTML = loc.aliases && loc.aliases.length > 0
                    ? `<p class="text-xs text-slate-500 mt-1">Apelidos: ${loc.aliases.join(', ')}</p>`
                    : '';
                tr.innerHTML = `
                    <td class="p-4">
                        <p class="font-semibold text-slate-700">${loc.name}</p>
                        ${aliasesHTML}
                    </td>
                    <td class="p-4 text-slate-600">${loc.code || 'N/A'}</td>
                    <td class="p-4 text-center">
                        <div class="flex justify-center items-center gap-1">
                            <button data-id="${loc.id}" class="edit-location-btn text-slate-500 hover:text-blue-600 p-2 rounded-full hover:bg-blue-100 transition" title="Editar"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg></button>
                            <button data-id="${loc.id}" class="delete-location-btn text-slate-500 hover:text-red-600 p-2 rounded-full hover:bg-red-100 transition" title="Excluir"><svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg></button>
                        </div>
                    </td>
                `;
                locationsList.appendChild(tr);
            });
        };

        const showAdjustModal = () => {
            const modalContent = `
                <h2 class="text-2xl font-semibold mb-6 text-center text-slate-800">Ajustar Estoque</h2>
                <form id="adjust-form" class="space-y-4">
                    <div>
                        <label class="block text-sm font-medium text-slate-600 mb-2">Tipo de Ajuste</label>
                        <div class="flex gap-4">
                            <label class="flex-1 flex items-center p-3 border-2 border-slate-200 rounded-lg cursor-pointer has-[:checked]:bg-indigo-50 has-[:checked]:border-indigo-400 transition-colors">
                                <input type="radio" name="adjust-type" value="entrada" class="mr-2 text-indigo-600 focus:ring-indigo-500" checked>
                                Entrada (Adicionar)
                            </label>
                            <label class="flex-1 flex items-center p-3 border-2 border-slate-200 rounded-lg cursor-pointer has-[:checked]:bg-indigo-50 has-[:checked]:border-indigo-400 transition-colors">
                                <input type="radio" name="adjust-type" value="saida" class="mr-2 text-indigo-600 focus:ring-indigo-500">
                                Saída (Remover)
                            </label>
                        </div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Quantidade</label>
                        <input type="number" id="quantity-change" class="mt-1 w-full p-3 border border-slate-200 rounded-lg text-center text-xl focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required min="1">
                    </div>
                    <div class="mt-8 flex justify-end space-x-4">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Confirmar Ajuste</button>
                    </div>
                </form>`;
            document.getElementById('generic-modal').innerHTML = modalContent;
            openModal('generic-modal');
        };

        const showRequisitionModal = (preselectedProductIds = new Set()) => {
            const sortedProducts = [...products].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            let productOptionsHTML = sortedProducts.map(p => `<option value="${p.id}">${p.name} (${p.codeRM || p.code}) - Estoque: ${p.quantity} - Local: ${p.location || 'N/A'}</option>`).join('');
            
            const content = `
                <h2 class="text-2xl font-semibold text-slate-800 mb-6">Criar Requisição</h2>
                <form id="requisition-form">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                        <div><label class="block text-sm font-medium">Nº Requisição</label><input type="text" id="req-number" class="w-full mt-1 p-2 border border-slate-200 bg-slate-100 rounded" value="${generateNextRequisitionNumber()}" readonly></div>
                        <div><label class="block text-sm font-medium">Data</label><input type="text" class="w-full mt-1 p-2 border border-slate-200 bg-slate-100 rounded" value="${new Date().toLocaleDateString('pt-BR')}" readonly></div>
                        <div><label class="block text-sm font-medium">Requisitante</label><input type="text" id="req-requester" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required></div>
                        <div><label class="block text-sm font-medium">Líder da Equipe</label><input type="text" id="req-team-leader" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required></div>
                        <div class="md:col-span-2">
                            <label class="block text-sm font-medium">Local de Aplicação</label>
                            <div class="relative">
                                <input type="text" id="req-location-search" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" placeholder="Pesquisar local ou apelido..." autocomplete="off">
                                <div id="req-location-results" class="absolute z-10 w-full bg-white border rounded mt-1 max-h-48 overflow-y-auto hidden shadow-lg">
                                </div>
                            </div>
                            <input type="hidden" id="req-application-location" name="req-application-location" required>
                        </div>
                        <div class="md:col-span-2">
                            <label class="block text-sm font-medium">Obra</label>
                            <select id="req-obra" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required>
                                <option value="TABOCA">TABOCA</option>
                                <option value="ESTRELA" selected>ESTRELA</option>
                            </select>
                        </div>
                    </div>
                    <h3 class="font-semibold mt-6 mb-2">Itens</h3>
                    <div class="mb-2">
                        <input type="text" id="req-item-search" class="w-full p-2 border border-slate-200 rounded" placeholder="Pesquisar item (nome, código, local, grupo)...">
                    </div>
                    <div id="req-items-container" class="space-y-2 max-h-60 overflow-y-auto p-2 bg-slate-50 border rounded-md"></div>
                    <button type="button" id="add-req-item-btn" class="mt-2 text-sm text-indigo-600 font-semibold hover:text-indigo-800">+ Adicionar Item</button>
                    <div class="mt-8 flex justify-end gap-4">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Salvar e Baixar Estoque</button>
                    </div>
                </form>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            document.getElementById('generic-modal').classList.replace('max-w-md', 'max-w-4xl');
            openModal('generic-modal');

            const reqLocationSearch = document.getElementById('req-location-search');
            const reqLocationResults = document.getElementById('req-location-results');
            const reqApplicationLocation = document.getElementById('req-application-location');
            const reqItemsContainer = document.getElementById('req-items-container');
            const reqItemSearch = document.getElementById('req-item-search');

            const renderLocationResults = (searchTerm = '') => {
                reqLocationResults.innerHTML = '';
                const lowerSearchTerm = searchTerm.toLowerCase();
                let filteredLocations = [];

                const allSortedLocations = [...locations].sort((a, b) => a.name.localeCompare(b.name));

                if (!searchTerm) {
                    filteredLocations = allSortedLocations.map(loc => ({...loc, match: loc.name}));
                } else {
                    allSortedLocations.forEach(loc => {
                        const lowerName = loc.name.toLowerCase();
                        if (lowerName.includes(lowerSearchTerm)) {
                            filteredLocations.push({...loc, match: loc.name});
                            return; 
                        }
                        if (loc.aliases) {
                            const matchingAlias = loc.aliases.find(alias => alias.toLowerCase().includes(lowerSearchTerm));
                            if (matchingAlias) {
                                filteredLocations.push({...loc, match: `Apelido: ${matchingAlias}`});
                            }
                        }
                    });
                }

                if (filteredLocations.length > 0) {
                    filteredLocations.forEach(loc => {
                        const resultItem = document.createElement('div');
                        resultItem.className = 'p-2 hover:bg-indigo-100 cursor-pointer';
                        resultItem.dataset.value = loc.name;
                        resultItem.innerHTML = `
                            <p class="font-semibold text-slate-800">${loc.name}</p>
                            ${loc.match !== loc.name ? `<p class="text-sm text-slate-500">${loc.match}</p>` : ''}
                        `;
                        resultItem.addEventListener('mousedown', () => {
                            reqLocationSearch.value = loc.name;
                            reqApplicationLocation.value = loc.name;
                            reqLocationResults.classList.add('hidden');
                        });
                        reqLocationResults.appendChild(resultItem);
                    });
                } else {
                    reqLocationResults.innerHTML = '<div class="p-2 text-slate-500">Nenhum local encontrado.</div>';
                }
                reqLocationResults.classList.remove('hidden');
            };

            reqLocationSearch.addEventListener('focus', () => renderLocationResults(reqLocationSearch.value));
            reqLocationSearch.addEventListener('input', () => renderLocationResults(reqLocationSearch.value));
            reqLocationSearch.addEventListener('blur', () => {
                setTimeout(() => {
                    reqLocationResults.classList.add('hidden');
                }, 150);
            });

            const normalizeSearchText = (value = '') => {
                return String(value)
                    .normalize('NFD')
                    .replace(/[\u0300-\u036f]/g, '')
                    .toLowerCase()
                    .trim();
            };

            const buildProductSearchText = (p) => {
                return [p.name, p.codeRM, p.code, p.location, p.group, p.unit].filter(Boolean).join(' ');
            };

            const getFilteredProducts = (rawSearchTerm = '') => {
                const normalizedTerm = normalizeSearchText(rawSearchTerm);
                if (!normalizedTerm) return sortedProducts;

                const tokens = normalizedTerm.split(/\s+/).filter(Boolean);

                return sortedProducts
                    .map((p) => {
                        const haystack = normalizeSearchText(buildProductSearchText(p));
                        if (!tokens.every(token => haystack.includes(token))) return null;

                        const normalizedName = normalizeSearchText(p.name || '');
                        const normalizedCodeRm = normalizeSearchText(p.codeRM || '');
                        const normalizedCode = normalizeSearchText(p.code || '');

                        let score = 0;
                        tokens.forEach(token => {
                            if (normalizedName.startsWith(token)) score += 8;
                            else if (normalizedName.includes(token)) score += 5;
                            if (normalizedCodeRm.startsWith(token) || normalizedCode.startsWith(token)) score += 4;
                        });

                        return { p, score };
                    })
                    .filter(Boolean)
                    .sort((a, b) => b.score - a.score || (a.p.name || '').localeCompare(b.p.name || ''))
                    .map(entry => entry.p);
            };
            
            reqItemSearch.addEventListener('input', () => {
                const filteredProducts = getFilteredProducts(reqItemSearch.value);

                reqItemsContainer.querySelectorAll('.req-item-product').forEach(select => {
                    const currentVal = select.value;
                    let newOptionsHTML = filteredProducts.map(p => `<option value="${p.id}">${p.name} (${p.codeRM || p.code}) - Estoque: ${p.quantity} - Local: ${p.location || 'N/A'}</option>`).join('');

                    if (currentVal && !filteredProducts.some(p => p.id === currentVal)) {
                        const currentProduct = sortedProducts.find(p => p.id === currentVal);
                        if (currentProduct) {
                            const currentOption = `<option value="${currentProduct.id}">${currentProduct.name} (${currentProduct.codeRM || currentProduct.code}) - Estoque: ${currentProduct.quantity} - Local: ${currentProduct.location || 'N/A'}</option>`;
                            newOptionsHTML = currentOption + newOptionsHTML;
                        }
                    }

                    if (!newOptionsHTML) {
                        newOptionsHTML = '<option value="">Nenhum item encontrado para esta busca</option>';
                    }
                    
                    select.innerHTML = newOptionsHTML;
                    select.value = currentVal;
                });
            });

            const addReqItem = (productIdToSelect = null) => {
                const itemRow = document.createElement('div');
                itemRow.className = 'flex items-center gap-2';
                const filteredProducts = getFilteredProducts(reqItemSearch.value);
                const optionsHTML = (filteredProducts.length > 0 ? filteredProducts : sortedProducts)
                    .map(p => `<option value="${p.id}">${p.name} (${p.codeRM || p.code}) - Estoque: ${p.quantity} - Local: ${p.location || 'N/A'}</option>`)
                    .join('');
                itemRow.innerHTML = `
                    <select class="req-item-product w-full p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500">${optionsHTML}</select>
                    <input type="number" class="req-item-quantity w-24 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" placeholder="Qtd." min="1" required>
                    <button type="button" class="remove-req-item-btn text-red-500 hover:text-red-700 font-bold p-1 text-2xl">&times;</button>
                `;
                reqItemsContainer.appendChild(itemRow);
                const newSelect = itemRow.querySelector('.req-item-product');
                if (productIdToSelect) {
                    newSelect.value = productIdToSelect;
                }
                newSelect.dispatchEvent(new Event('change'));
            };

            document.getElementById('add-req-item-btn').onclick = () => addReqItem();

            reqItemsContainer.addEventListener('click', e => {
                if (e.target.classList.contains('remove-req-item-btn')) {
                    e.target.parentElement.remove();
                }
            });
            
            reqItemsContainer.addEventListener('change', e => {
                if (e.target.classList.contains('req-item-product')) {
                    const selectedProductId = e.target.value;
                    const product = products.find(p => p.id === selectedProductId);
                    const quantityInput = e.target.nextElementSibling;
                    if (product && quantityInput) {
                        quantityInput.max = product.quantity;
                        quantityInput.placeholder = `Max: ${product.quantity}`;
                    } else if (quantityInput) {
                        quantityInput.removeAttribute('max');
                        quantityInput.placeholder = 'Qtd.';
                    }
                }
            });

            if (preselectedProductIds.size > 0) {
                preselectedProductIds.forEach(id => {
                    addReqItem(id);
                });
            } else {
                addReqItem();
            }
        };

        const showViewRequisitionModal = (reqIds) => {
            const reqs = requisitions.filter(r => reqIds.includes(r.id));
            if (reqs.length === 0) return;

            let allContentHTML = '';
            reqs.forEach((req, index) => {
                let itemsHTML = req.items.map(item => `
                    <tr>
                        <td class="border p-2">${item.productCodeRM || item.productCode}</td>
                        <td class="border p-2">${item.productName}</td>
                        <td class="border p-2 text-center">${item.quantity}</td>
                    </tr>
                `).join('');
                
                const logoHTML = appSettings.logoUrl ? `<img src="${appSettings.logoUrl}" class="h-12 w-auto max-w-[150px] object-contain">` : '';
                const pageBreak = index < reqs.length - 1 ? 'page-break-after: always;' : '';

                allContentHTML += `
                    <div style="${pageBreak}">
                        <div class="flex items-center justify-between mb-6 border-b pb-4">
                            <div class="flex items-center gap-4">
                                ${logoHTML}
                                <h1 class="text-xl font-bold text-slate-800">${appSettings.appName}</h1>
                            </div>
                            <div class="text-right">
                                <h2 class="text-lg font-bold">Requisição de Material</h2>
                                <p class="text-sm">Nº ${req.number}</p>
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-x-4 gap-y-2 mb-6 text-sm">
                            <p><strong>Data:</strong> ${req.date ? new Date(req.date.seconds * 1000).toLocaleDateString('pt-BR') : '...'}</p>
                            <p><strong>Requisitante:</strong> ${req.requester}</p>
                            <p><strong>Líder da Equipe:</strong> ${req.teamLeader || 'N/A'}</p>
                            <p><strong>Obra:</strong> ${req.obra || 'N/A'}</p>
                            <p class="col-span-2"><strong>Local de Aplicação:</strong> ${req.applicationLocation}</p>
                        </div>
                        <table class="w-full text-left border-collapse mt-4 text-sm">
                            <thead class="bg-slate-100">
                                <tr>
                                    <th class="border p-2">Código</th>
                                    <th class="border p-2">Produto</th>
                                    <th class="border p-2 text-center">Quantidade</th>
                                </tr>
                            </thead>
                            <tbody>${itemsHTML}</tbody>
                        </table>
                         <div class="mt-16 text-sm">
                              <div class="flex justify-around items-end">
                                  <div class="text-center">
                                      <p class="border-t border-black pt-1 px-12">Assinatura Requisitante</p>
                                  </div>
                                  <div class="text-center">
                                      <p class="border-t border-black pt-1 px-12">Assinatura Líder</p>
                                  </div>
                                  <div class="text-center">
                                      <p class="border-t border-black pt-1 px-12">Assinatura Almoxarifado</p>
                                  </div>
                              </div>
                         </div>
                    </div>
                `;
            });

            const content = `
                <div id="pdf-content" class="bg-white p-4">${allContentHTML}</div>
                <div class="mt-8 flex justify-end gap-4">
                    <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Fechar</button>
                    <button type="button" id="download-pdf-btn" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Baixar PDF</button>
                </div>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            document.getElementById('generic-modal').classList.replace('max-w-md', 'max-w-4xl');
            openModal('generic-modal');

            document.getElementById('download-pdf-btn').onclick = () => {
                const { jsPDF } = window.jspdf;
                const requisitionElement = document.getElementById('pdf-content');
                
                const btn = document.getElementById('download-pdf-btn');
                btn.innerHTML = `<div class="spinner-small"></div>`;
                btn.disabled = true;

                html2canvas(requisitionElement, { scale: 2 }).then(canvas => {
                    const imgData = canvas.toDataURL('image/png');
                    const pdf = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
                    const pdfWidth = pdf.internal.pageSize.getWidth();
                    const canvasWidth = canvas.width;
                    const canvasHeight = canvas.height;
                    const ratio = canvasWidth / canvasHeight;
                    const width = pdfWidth - 40;
                    const height = width / ratio;

                    pdf.addImage(imgData, 'PNG', 20, 20, width, height);
                    pdf.save(`Requisicoes.pdf`);
                    
                    btn.innerHTML = `Baixar PDF`;
                    btn.disabled = false;
                });
            };
        };

        const showEditModal = () => {
            const p = products.find(prod => prod.id === currentProductId);
            if (!p) return;

            const units = ["Unidade", "Peça", "Metros", "Litro", "Quilo", "Caixa", "Pacote", "Rolo", "Galão", "Saco"];
            const unitOptions = units.map(unit => `<option value="${unit}" ${p.unit === unit ? 'selected' : ''}>${unit}</option>`).join('');
            
            const groups = ["Elétrico", "Hidráulico", "Consumível", "Ferramentas Manuais", "Material de Corte e Solda", "Escritório", "Segurança", "Outros"];
            const groupOptions = groups.map(group => `<option value="${group}" ${p.group === group ? 'selected' : ''}>${group}</option>`).join('');

            const content = `
                <h2 class="text-2xl font-semibold mb-6 text-slate-800">Editar Produto</h2>
                <form id="edit-product-form" class="space-y-4">
                    <input type="hidden" id="edit-product-id" value="${p.id}">
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div><label class="block text-sm font-medium text-slate-600">Nome</label><input type="text" id="edit-product-name" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${p.name}" required></div>
                        <div><label class="block text-sm font-medium text-slate-600">Código RM (Externo)</label><input type="text" id="edit-product-code-rm" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${p.codeRM || ''}" required></div>
                        <div><label class="block text-sm font-medium text-slate-600">Código Interno (SKU)</label><input type="text" id="edit-product-code" class="w-full mt-1 p-3 border border-slate-200 rounded-lg bg-slate-100" value="${p.code}" readonly></div>
                        <div><label class="block text-sm font-medium text-slate-600">Localização</label><input type="text" id="edit-product-location" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${p.location}" required></div>
                        <div><label class="block text-sm font-medium text-slate-600">Grupo</label><select id="edit-product-group" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500">${groupOptions}</select></div>
                        <div><label class="block text-sm font-medium text-slate-600">Unidade</label><select id="edit-product-unit" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500">${unitOptions}</select></div>
                        <div><label class="block text-sm font-medium text-slate-600">Quantidade Atual</label><input type="number" id="edit-product-quantity" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${p.quantity}" required min="0"></div>
                        <div><label class="block text-sm font-medium text-slate-600">Qtd. Mínima</label><input type="number" id="edit-product-min-quantity" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${p.minQuantity}" required min="0"></div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Observação</label>
                        <textarea id="edit-product-observation" class="w-full mt-1 p-3 border border-slate-200 rounded-lg h-24 focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500">${p.observation || ''}</textarea>
                    </div>
                    <div class="mt-8 flex justify-end space-x-4">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Salvar Alterações</button>
                    </div>
                </form>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            openModal('generic-modal');
        };

        const showHistoryModal = () => {
             const product = products.find(p => p.id === currentProductId);
            if (!product) return;

            const productHistory = history
                .filter(h => h.productId === currentProductId)
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            let historyItemsHTML = '';
            if (productHistory.length > 0) {
                productHistory.forEach(h => {
                    let details = '';
                    let colorClass = '';
                    switch(h.type) {
                        case 'Entrada': case 'Ajuste Entrada':
                            colorClass = 'border-green-500'; details = `<strong>+${h.quantity}</strong> unidades. Estoque foi para <strong>${h.newTotal}</strong>.`; break;
                        case 'Saída': case 'Saída por Requisição':
                            colorClass = 'border-red-500'; details = `<strong>-${h.quantity}</strong> unidades. Retirado por <strong>${h.withdrawnBy}</strong>. Estoque foi para <strong>${h.newTotal}</strong>.`; break;
                        case 'Ajuste Saída':
                            colorClass = 'border-yellow-500'; details = `<strong>-${h.quantity}</strong> unidades ajustadas. Estoque foi para <strong>${h.newTotal}</strong>.`; break;
                        case 'Criação': case 'Importação': 
                            colorClass = 'border-blue-500'; details = `Produto cadastrado com <strong>${h.quantity}</strong> unidades.`; break;
                        case 'Edição': 
                            colorClass = 'border-purple-500'; details = `Dados do produto foram alterados.`; break;
                    }
                    historyItemsHTML += `
                        <div class="p-3 bg-slate-50 rounded-lg border-l-4 ${colorClass}">
                            <p class="font-semibold text-slate-800">${h.type}</p>
                            <p class="text-sm text-slate-700">${details}</p>
                            <p class="text-xs text-slate-500 mt-1">${h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : 'Data indisponível'}</p>
                        </div>
                    `;
                });
            } else {
                historyItemsHTML = '<p class="text-slate-500 text-center">Nenhum histórico para este produto.</p>';
            }

            const content = `
                <div class="flex justify-between items-start">
                    <div>
                        <h2 class="text-2xl font-bold">Histórico do Produto</h2>
                        <p class="text-slate-600 mb-6">${product.name} (${product.codeRM || product.code})</p>
                    </div>
                    <button type="button" class="close-modal-btn p-2 -mt-2 -mr-2 text-slate-400 hover:text-slate-700 text-3xl">&times;</button>
                </div>
                <div class="max-h-96 overflow-y-auto space-y-3 pr-2">${historyItemsHTML}</div>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            document.getElementById('generic-modal').classList.replace('max-w-md', 'max-w-4xl');
            openModal('generic-modal');
        };
        
        const showSettingsModal = () => {
            const logoPreviewSrc = appSettings.logoUrl || 'https://placehold.co/200x60/e2e8f0/475569?text=Sem+Logo';
            const currentZoom = parseInt(localStorage.getItem('appZoomLevel') || '100');
            const content = `
                <h2 class="text-2xl font-semibold mb-6 text-slate-800">Configurações</h2>
                <form id="settings-form" class="space-y-4">
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Nome do App</label>
                        <input type="text" id="settings-app-name" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${appSettings.appName}" required>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600 mb-2">Logo</label>
                        <div class="flex items-center gap-4 p-3 border border-slate-200 rounded-lg bg-slate-50">
                            <img id="logo-preview" src="${logoPreviewSrc}" class="w-32 h-12 object-contain rounded-md bg-white border">
                            <div class="flex-grow space-y-2">
                                <button type="button" id="upload-logo-btn" class="w-full text-sm px-4 py-2 bg-white border border-slate-300 rounded-lg font-semibold hover:bg-slate-100 transition">Carregar Imagem</button>
                                <button type="button" id="remove-logo-btn" class="w-full text-sm px-4 py-2 text-red-600 font-semibold hover:bg-red-50 transition">Remover Logo</button>
                            </div>
                            <input type="file" id="logo-upload-input" class="hidden" accept="image/*">
                        </div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600 mb-2">
                            <span class="material-symbols-outlined align-middle mr-1" style="font-size:18px;">zoom_in</span>
                            Ajuste de Tela / Zoom
                        </label>
                        <p class="text-xs text-slate-400 mb-3">Ajuste o tamanho da interface para melhor visualização no seu dispositivo.</p>
                        <div class="flex items-center gap-3 p-3 border border-slate-200 rounded-lg bg-slate-50">
                            <button type="button" id="zoom-decrease-btn" class="w-9 h-9 flex items-center justify-center rounded-lg bg-white border border-slate-300 font-bold text-lg text-slate-700 hover:bg-slate-100 transition flex-shrink-0">−</button>
                            <div class="flex-grow">
                                <input type="range" id="zoom-range" min="70" max="130" step="5" value="${currentZoom}" class="w-full accent-blue-600" style="cursor:pointer;">
                            </div>
                            <button type="button" id="zoom-increase-btn" class="w-9 h-9 flex items-center justify-center rounded-lg bg-white border border-slate-300 font-bold text-lg text-slate-700 hover:bg-slate-100 transition flex-shrink-0">+</button>
                            <span id="zoom-value" class="text-sm font-bold text-slate-700 min-w-[3rem] text-center">${currentZoom}%</span>
                        </div>
                        <div class="flex gap-2 mt-2">
                            <button type="button" id="zoom-preset-85" class="text-xs px-2.5 py-1 rounded-full border font-semibold transition ${currentZoom === 85 ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}">85%</button>
                            <button type="button" id="zoom-preset-90" class="text-xs px-2.5 py-1 rounded-full border font-semibold transition ${currentZoom === 90 ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}">90%</button>
                            <button type="button" id="zoom-preset-100" class="text-xs px-2.5 py-1 rounded-full border font-semibold transition ${currentZoom === 100 ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}">100%</button>
                            <button type="button" id="zoom-preset-110" class="text-xs px-2.5 py-1 rounded-full border font-semibold transition ${currentZoom === 110 ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}">110%</button>
                            <button type="button" id="zoom-preset-120" class="text-xs px-2.5 py-1 rounded-full border font-semibold transition ${currentZoom === 120 ? 'bg-blue-600 text-white border-blue-600' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'}">120%</button>
                        </div>
                    </div>
                    <div class="mt-8 flex justify-end space-x-4">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Salvar</button>
                    </div>
                </form>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            openModal('generic-modal');

            // Zoom control logic
            const zoomRange = document.getElementById('zoom-range');
            const zoomValue = document.getElementById('zoom-value');
            const zoomDecreaseBtn = document.getElementById('zoom-decrease-btn');
            const zoomIncreaseBtn = document.getElementById('zoom-increase-btn');

            const applyZoomPreview = (val) => {
                zoomRange.value = val;
                zoomValue.textContent = val + '%';
                applyZoom(val);
                // Update preset button styles
                [85, 90, 100, 110, 120].forEach(p => {
                    const btn = document.getElementById('zoom-preset-' + p);
                    if (btn) {
                        if (parseInt(val) === p) {
                            btn.className = 'text-xs px-2.5 py-1 rounded-full border font-semibold transition bg-blue-600 text-white border-blue-600';
                        } else {
                            btn.className = 'text-xs px-2.5 py-1 rounded-full border font-semibold transition bg-white text-slate-600 border-slate-300 hover:bg-slate-100';
                        }
                    }
                });
            };

            zoomRange.addEventListener('input', (e) => applyZoomPreview(e.target.value));
            zoomDecreaseBtn.addEventListener('click', () => {
                const newVal = Math.max(70, parseInt(zoomRange.value) - 5);
                applyZoomPreview(newVal);
            });
            zoomIncreaseBtn.addEventListener('click', () => {
                const newVal = Math.min(130, parseInt(zoomRange.value) + 5);
                applyZoomPreview(newVal);
            });
            [85, 90, 100, 110, 120].forEach(preset => {
                const btn = document.getElementById('zoom-preset-' + preset);
                if (btn) btn.addEventListener('click', () => applyZoomPreview(preset));
            });
        };

        // --- Zoom / Scale Functions ---
        const applyZoom = (level) => {
            const val = parseInt(level) || 100;
            document.documentElement.style.fontSize = (val / 100) * 16 + 'px';
            localStorage.setItem('appZoomLevel', val);
        };

        // Apply saved zoom on page load
        (() => {
            const savedZoom = localStorage.getItem('appZoomLevel');
            if (savedZoom && savedZoom !== '100') {
                applyZoom(savedZoom);
            }
        })();
        
        const showEditLocationModal = () => {
            const loc = locations.find(l => l.id === currentLocationId);
            if (!loc) return;
            const aliasesStr = (loc.aliases || []).join(', ');
            const content = `
                <h2 class="text-2xl font-semibold mb-6 text-slate-800">Editar Local</h2>
                <form id="edit-location-form" class="space-y-4">
                    <input type="hidden" id="edit-location-id" value="${loc.id}">
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Nome do Local</label>
                        <input type="text" id="edit-location-name" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${loc.name}" required>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Código (Opcional)</label>
                        <input type="text" id="edit-location-code" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${loc.code || ''}">
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-600">Apelidos (separados por vírgula)</label>
                        <input type="text" id="edit-location-aliases" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" value="${aliasesStr}">
                    </div>
                    <div class="mt-8 flex justify-end space-x-4">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Salvar Alterações</button>
                    </div>
                </form>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            openModal('generic-modal');
        };

        const renderEntryView = () => {
            // Renderização da busca de produtos será feita via listeners de input
        };

        const setupEntryViewListeners = () => {
            const entryProductSearch = document.getElementById('entry-product-search');
            const entryProductsContainer = document.getElementById('entry-products-container');
            if (!entryProductSearch || !entryProductsContainer) return;

            // Limpar listeners antigos (se existirem)
            const newSearch = entryProductSearch.cloneNode(true);
            entryProductSearch.replaceWith(newSearch);
            const freshSearch = document.getElementById('entry-product-search');

            const sortedProducts = [...products].sort((a, b) => a.name.localeCompare(b.name));

            const renderProductOptions = (searchTerm = '') => {
                entryProductsContainer.innerHTML = '';
                const normalizedTerm = normalizeSearchText(searchTerm);
                let filteredProducts = sortedProducts;

                if (normalizedTerm) {
                    const tokens = normalizedTerm.split(/\s+/).filter(Boolean);
                    filteredProducts = sortedProducts
                        .map((p) => {
                            const haystack = normalizeSearchText(buildProductSearchText(p));
                            if (!tokens.every(token => haystack.includes(token))) return null;

                            const normalizedName = normalizeSearchText(p.name || '');
                            const normalizedCodeRm = normalizeSearchText(p.codeRM || '');
                            const normalizedCode = normalizeSearchText(p.code || '');

                            let score = 0;
                            tokens.forEach(token => {
                                if (normalizedName.startsWith(token)) score += 8;
                                else if (normalizedName.includes(token)) score += 5;
                                if (normalizedCodeRm.startsWith(token) || normalizedCode.startsWith(token)) score += 4;
                            });

                            return { p, score };
                        })
                        .filter(Boolean)
                        .sort((a, b) => b.score - a.score || (a.p.name || '').localeCompare(b.p.name || ''))
                        .map(entry => entry.p);
                }

                if (filteredProducts.length === 0) {
                    entryProductsContainer.innerHTML = '<div class="p-3 text-slate-500 text-center">Nenhum produto encontrado</div>';
                    entryProductsContainer.classList.remove('hidden');
                    return;
                }

                filteredProducts.forEach(p => {
                    const resultItem = document.createElement('div');
                    resultItem.className = 'p-3 hover:bg-green-100 cursor-pointer border-b border-slate-200 last:border-b-0 transition-colors';
                    resultItem.innerHTML = `
                        <p class="font-semibold text-slate-800">${p.name}</p>
                        <p class="text-sm text-slate-500">Cód. RM: ${p.codeRM || p.code} | Estoque: ${p.quantity} | Local: ${p.location || 'N/A'}</p>
                    `;
                    resultItem.addEventListener('mousedown', () => {
                        document.getElementById('entry-product-select').value = p.id;
                        freshSearch.value = p.name;
                        entryProductsContainer.classList.add('hidden');
                    });
                    entryProductsContainer.appendChild(resultItem);
                });

                entryProductsContainer.classList.remove('hidden');
            };

            freshSearch.addEventListener('focus', () => {
                renderProductOptions(freshSearch.value);
            });

            freshSearch.addEventListener('input', () => {
                renderProductOptions(freshSearch.value);
            });

            // Fechar container quando clicar fora
            const closeContainer = (e) => {
                if (freshSearch && entryProductsContainer &&
                    !freshSearch.contains(e.target) && !entryProductsContainer.contains(e.target)) {
                    entryProductsContainer.classList.add('hidden');
                }
            };
            document.addEventListener('mousedown', closeContainer);
        };


        const renderExitView = () => {
            const productSelect = document.getElementById('exit-product-select');
            const locationDatalist = document.getElementById('location-datalist');
            
            const sortedProducts = [...products].sort((a, b) => a.name.localeCompare(b.name));
            productSelect.innerHTML = '<option value="">Selecione um produto...</option>';
            sortedProducts.forEach(p => {
                const option = document.createElement('option');
                option.value = p.id;
                option.textContent = `${p.name} (${p.codeRM || p.code}) - Estoque: ${p.quantity}`;
                productSelect.appendChild(option);
            });

            const sortedLocations = [...locations].sort((a, b) => a.name.localeCompare(b.name));
            locationDatalist.innerHTML = '';
            sortedLocations.forEach(l => {
                const option = document.createElement('option');
                option.value = l.name;
                option.textContent = l.code ? `${l.name} (${l.code})` : l.name;
                locationDatalist.appendChild(option);
            });
        };

        const renderComprasView = (periodDays = 15) => {
            const now = Date.now();
            const periodMs = periodDays * 24 * 60 * 60 * 1000;
            const cutoff = now - periodMs;

            // Calcular consumo por produto no período
            const consumoMap = {};
            history.forEach(h => {
                if (!h.productId) return;
                const isExit = h.type === 'Saída' || h.type === 'Saída por Requisição' || h.type === 'Ajuste Saída';
                if (!isExit) return;
                const ts = h.date?.seconds ? h.date.seconds * 1000 : 0;
                if (ts < cutoff) return;
                consumoMap[h.productId] = (consumoMap[h.productId] || 0) + (h.quantity || 0);
            });

            const activeFilter = document.querySelector('.compras-filter-btn[style*="background:#191c1d"]')?.dataset.filter || 'all';

            // Montar lista com dados calculados
            const items = products.map(p => {
                const totalConsumed = consumoMap[p.id] || 0;
                const dailyRate = totalConsumed / periodDays;
                const daysRemaining = dailyRate > 0 ? Math.floor(p.quantity / dailyRate) : null;
                const repoQty30days = Math.ceil(dailyRate * 30 * 1.2); // 30 dias + 20% margem
                const suggestedQty = Math.max(0, repoQty30days - p.quantity);

                let status;
                if (p.quantity <= 0) status = 'critico';
                else if (p.quantity <= p.minQuantity) status = 'critico';
                else if (daysRemaining !== null && daysRemaining <= 15) status = 'atencao';
                else if (p.quantity <= p.minQuantity * 1.5) status = 'atencao';
                else if (totalConsumed > 0) status = 'monitorar';
                else return null; // não consumido e ok → sem relevância

                return { p, totalConsumed, dailyRate, daysRemaining, suggestedQty, status };
            }).filter(Boolean);

            // Ordenar: critico → atencao → monitorar, depois por dias restantes
            const ordem = { critico: 0, atencao: 1, monitorar: 2 };
            items.sort((a, b) => {
                if (ordem[a.status] !== ordem[b.status]) return ordem[a.status] - ordem[b.status];
                const dA = a.daysRemaining ?? 9999;
                const dB = b.daysRemaining ?? 9999;
                return dA - dB;
            });

            // Cards de resumo
            const criticos = items.filter(i => i.status === 'critico').length;
            const atencoes = items.filter(i => i.status === 'atencao').length;
            const monitorar = items.filter(i => i.status === 'monitorar').length;
            const totalSugerido = items.reduce((acc, i) => acc + i.suggestedQty, 0);

            const summaryEl = document.getElementById('compras-summary');
            summaryEl.innerHTML = `
                <div class="rounded-2xl p-4" style="background:#fff2f0; border:1px solid #f8bdb4;">
                    <p class="text-xs font-bold uppercase tracking-wider" style="color:#ba1a1a;">🔴 Crítico</p>
                    <p class="text-3xl font-black mt-1" style="color:#ba1a1a;">${criticos}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">itens no limite</p>
                </div>
                <div class="rounded-2xl p-4" style="background:#fff8e1; border:1px solid #f9e099;">
                    <p class="text-xs font-bold uppercase tracking-wider" style="color:#795900;">🟡 Atenção</p>
                    <p class="text-3xl font-black mt-1" style="color:#795900;">${atencoes}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">comprar em breve</p>
                </div>
                <div class="rounded-2xl p-4" style="background:#f3f4f5; border:1px solid #c1c6d6;">
                    <p class="text-xs font-bold uppercase tracking-wider" style="color:#414754;">🔵 Monitorar</p>
                    <p class="text-3xl font-black mt-1" style="color:#414754;">${monitorar}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">consumo ativo</p>
                </div>
                <div class="rounded-2xl p-4" style="background:#e8f5e9; border:1px solid #a5d6a7;">
                    <p class="text-xs font-bold uppercase tracking-wider" style="color:#006e2c;">📦 Itens p/ Comprar</p>
                    <p class="text-3xl font-black mt-1" style="color:#006e2c;">${criticos + atencoes}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">compra sugerida: ${totalSugerido} unid.</p>
                </div>
            `;

            // Filtrar por status selecionado
            const filtered = activeFilter === 'all' ? items : items.filter(i => i.status === activeFilter);
            lastComprasFilteredItems = filtered;
            lastComprasPeriodDays = periodDays;

            const tbody = document.getElementById('compras-list');
            const noMsg = document.getElementById('no-compras-message');
            tbody.innerHTML = '';

            if (filtered.length === 0) {
                noMsg.classList.remove('hidden');
                return;
            }
            noMsg.classList.add('hidden');

            filtered.forEach(({ p, totalConsumed, dailyRate, daysRemaining, suggestedQty, status }) => {
                let badgeBg, badgeColor, badgeLabel;
                if (status === 'critico') { badgeBg = '#fff2f0'; badgeColor = '#ba1a1a'; badgeLabel = '🔴 CRÍTICO'; }
                else if (status === 'atencao') { badgeBg = '#fff8e1'; badgeColor = '#795900'; badgeLabel = '🟡 ATENÇÃO'; }
                else { badgeBg = '#f3f4f5'; badgeColor = '#414754'; badgeLabel = '🔵 MONITORAR'; }

                const daysText = daysRemaining !== null
                    ? (daysRemaining === 0 ? '<span style="color:#ba1a1a; font-weight:800;">ESGOTADO</span>'
                        : `<span style="color:${daysRemaining <= 7 ? '#ba1a1a' : daysRemaining <= 15 ? '#795900' : '#006e2c'}; font-weight:700;">${daysRemaining}d</span>`)
                    : '<span style="color:#c1c6d6;">—</span>';

                const dailyText = dailyRate > 0 ? dailyRate.toFixed(1) : '<span style="color:#c1c6d6;">—</span>';

                const tr = document.createElement('tr');
                tr.className = 'transition-colors hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="px-4 py-3">
                        <span class="px-2 py-1 rounded-full text-xs font-bold" style="background:${badgeBg}; color:${badgeColor};">${badgeLabel}</span>
                    </td>
                    <td class="px-4 py-3">
                        <p class="font-bold text-slate-800" style="font-size:13px;">${p.name}</p>
                        <p class="text-xs text-slate-400">${p.codeRM || p.code} · ${p.location || 'N/A'}</p>
                    </td>
                    <td class="px-4 py-3 text-center font-bold" style="color:${p.quantity <= p.minQuantity ? '#ba1a1a' : '#191c1d'};">${p.quantity} <span class="text-xs font-normal text-slate-400">${p.unit || ''}</span></td>
                    <td class="px-4 py-3 text-center text-slate-500">${p.minQuantity}</td>
                    <td class="px-4 py-3 text-center text-slate-600">${dailyText}</td>
                    <td class="px-4 py-3 text-center">${daysText}</td>
                    <td class="px-4 py-3 text-center">
                        ${suggestedQty > 0
                            ? `<span class="px-3 py-1 rounded-full text-xs font-black" style="background:#e8f5e9; color:#006e2c;">${suggestedQty} ${p.unit || 'un'}</span>`
                            : '<span style="color:#c1c6d6;">—</span>'}
                    </td>
                `;
                tbody.appendChild(tr);
            });
        };

        const switchView = (viewId) => {
            currentViewId = viewId;
            if (viewId !== 'inventory-view' && viewId !== 'requisitions-view') {
                selectedProductIds.clear();
            }
            views.forEach(view => view.classList.add('hidden'));
            document.getElementById(viewId).classList.remove('hidden');
            tabButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.view === viewId));
            
            // Sync bottom-nav active state
            document.querySelectorAll('.bottom-nav-btn').forEach(btn => {
                const isActive = btn.dataset.view === viewId;
btn.style.color = isActive ? '#0066FF' : '#6b7280';
                btn.style.background = isActive ? 'rgba(0,102,255,0.08)' : 'transparent';
            });

            if (viewId === 'add-product-view') {
                document.getElementById('product-code').value = generateNextProductCode();
            }
            if (viewId === 'entry-view') {
                renderEntryView();
                setupEntryViewListeners();
            }
            if (viewId === 'exit-log-view') renderExitLog(exitsSearchInput.value);
            if (viewId === 'rm-view') renderRMView(rmSearchInput.value); 
            if (viewId === 'dashboard-view') updateDashboard();
            if (viewId === 'exit-view') renderExitView();
            if (viewId === 'compras-view') {
                const sel = document.getElementById('compras-period-select');
                renderComprasView(sel ? parseInt(sel.value) : 15);
            }
            if (viewId === 'estrela-view') renderEstrelaView();

            syncRealtimeListeners();
        };

        const updateSelectionActionButtonsState = () => {
            const count = selectedProductIds.size;
            const initiateRequisitionBtn = document.getElementById('initiate-requisition-btn');
            
            const deleteBtnSpan = deleteSelectedBtn.querySelector('span');
            const reqBtnSpan = initiateRequisitionBtn.querySelector('span');

            if (count > 0) {
                deleteSelectedBtn.classList.remove('hidden');
                if(deleteBtnSpan) deleteBtnSpan.textContent = `Excluir (${count})`;
                initiateRequisitionBtn.classList.remove('hidden');
                if(reqBtnSpan) reqBtnSpan.textContent = `Iniciar Req. (${count})`;
            } else {
                deleteSelectedBtn.classList.add('hidden');
                initiateRequisitionBtn.classList.add('hidden');
            }

            // Botão de IA desativado (sem API Key configurada)
            aiDescribeBtn.classList.add('hidden');

            const allCheckboxes = document.querySelectorAll('.product-checkbox');
            const selectAllCheckbox = document.getElementById('select-all-products');
            if (!selectAllCheckbox) return;

            const displayedCheckedCount = document.querySelectorAll('.product-checkbox:checked').length;

            if (displayedCheckedCount === 0) {
                selectAllCheckbox.checked = false;
                selectAllCheckbox.indeterminate = selectedProductIds.size > 0;
            } else if (displayedCheckedCount === allCheckboxes.length && allCheckboxes.length > 0) {
                selectAllCheckbox.checked = true;
                selectAllCheckbox.indeterminate = false;
            } else {
                selectAllCheckbox.indeterminate = true;
            }
        };

        // --- Funções de Utilitários ---
        const copyToClipboard = (text) => {
            if (!text) {
                showToast("Nada para copiar.", true);
                return;
            }
            const textArea = document.createElement("textarea");
            textArea.value = text;
            document.body.appendChild(textArea);
            textArea.select();
            try {
                document.execCommand('copy');
                showToast(`"${text}" copiado!`);
            } catch (err) {
                showToast('Falha ao copiar.', true);
            }
            document.body.removeChild(textArea);
        };

        // --- Speech Recognition ---
        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
        let recognition;
        if (SpeechRecognition) {
            recognition = new SpeechRecognition();
            recognition.continuous = false;
            recognition.lang = 'pt-BR';
            recognition.interimResults = false;
            recognition.maxAlternatives = 1;

            recognition.onresult = (event) => {
                const transcript = event.results[0][0].transcript;
                aiExitPrompt.value = transcript;
                aiExitPrompt.dispatchEvent(new Event('change'));
            };

            recognition.onspeechend = () => {
                recognition.stop();
                micBtn.classList.remove('is-listening');
                document.getElementById('mic-status').textContent = 'Clique no microfone para falar';
            };

            recognition.onerror = (event) => {
                micBtn.classList.remove('is-listening');
                document.getElementById('mic-status').textContent = 'Clique no microfone para falar';
                if (event.error !== 'no-speech') {
                    showToast(`Erro no reconhecimento de voz: ${event.error}`, true);
                }
            };
        } else {
            micBtn.style.display = 'none';
            document.getElementById('mic-status').textContent = 'Reconhecimento de voz não suportado.';
        }


        // --- Bottom Nav + Drawer ---
        const openDrawer = () => {
            document.getElementById('bottom-drawer').classList.remove('translate-y-full');
            document.getElementById('bottom-drawer-backdrop').classList.remove('hidden');
        };
        const closeDrawer = () => {
            document.getElementById('bottom-drawer').classList.add('translate-y-full');
            document.getElementById('bottom-drawer-backdrop').classList.add('hidden');
        };

        document.querySelectorAll('.bottom-nav-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const view = btn.dataset.view;
                if (view) switchView(view);
            });
        });

        document.querySelectorAll('.drawer-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const view = btn.dataset.view;
                if (view) { switchView(view); closeDrawer(); }
            });
        });

        const moreBtn = document.getElementById('bottom-nav-more-btn');
        if (moreBtn) moreBtn.addEventListener('click', openDrawer);
        const drawerBackdrop = document.getElementById('bottom-drawer-backdrop');
        if (drawerBackdrop) drawerBackdrop.addEventListener('click', closeDrawer);

        // --- Event Listeners ---

        // 🔐 Login com usuário e senha (sem Firebase Auth)
        const doLogin = async () => {
            const username = loginUsernameInput.value.trim().toUpperCase();
            const password = loginPasswordInput.value;
            loginError.classList.add('hidden');

            if (!username || !password) {
                loginError.textContent = 'Preencha o nome de usuário e a senha.';
                loginError.classList.remove('hidden');
                return;
            }

            loginBtn.disabled = true;
            loginBtn.textContent = 'Verificando...';

            const hash = await hashPassword(password);
            const userId = normalizeUserId(username);

            // 1. Verificar credenciais hardcoded (funciona sem internet)
            const BUILTIN_USERS = {
                'uhe_estrela': { displayName: 'UHE ESTRELA', role: 'admin', pwd: '60218' }
            };
            const builtin = BUILTIN_USERS[userId];
            if (builtin) {
                const expectedHash = await hashPassword(builtin.pwd);
                if (hash === expectedHash) {
                    const appUser = { uid: userId, displayName: builtin.displayName, role: builtin.role };
                    localStorage.setItem('appUser', JSON.stringify(appUser));
                    initializeAppSession(appUser);
                    return;
                } else {
                    loginError.textContent = 'Senha incorreta.';
                    loginError.classList.remove('hidden');
                    loginBtn.disabled = false;
                    loginBtn.textContent = 'Entrar';
                    return;
                }
            }

            // 2. Verificar no Firestore para outros usuários
            try {
                const userRef = doc(db, `/artifacts/${appId}/public/data/users`, userId);
                const userDoc = await getDoc(userRef);
                if (userDoc.exists() && userDoc.data().passwordHash === hash) {
                    const data = userDoc.data();
                    const appUser = { uid: userId, displayName: data.displayName || data.username || username, role: data.role || 'operador' };
                    localStorage.setItem('appUser', JSON.stringify(appUser));
                    initializeAppSession(appUser);
                } else {
                    loginError.textContent = userDoc.exists() ? 'Senha incorreta.' : 'Usuário não encontrado.';
                    loginError.classList.remove('hidden');
                    loginBtn.disabled = false;
                    loginBtn.textContent = 'Entrar';
                }
            } catch (e) {
                console.error('Erro ao verificar usuário no Firestore:', e);
                loginError.textContent = 'Usuário não encontrado.';
                loginError.classList.remove('hidden');
                loginBtn.disabled = false;
                loginBtn.textContent = 'Entrar';
            }
        };

        loginBtn.addEventListener('click', doLogin);
        loginPasswordInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') doLogin(); });
        loginUsernameInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') loginPasswordInput.focus(); });

        logoutBtn.addEventListener('click', () => {
            stopEstrelaListeners();
            stopCoreListeners();
            localStorage.removeItem('appUser');
            currentUser = null;
            userRole = null;
            isDataLoaded = false;
            isAuthInitialized = false;
            products = [];
            history = [];
            requisitions = [];
            locations = [];
            loginScreen.classList.remove('hidden');
            appContainer.classList.add('hidden');
            if (logoutBtn) logoutBtn.classList.add('hidden');
            showToast('Logout realizado com sucesso');
        });

        settingsBtn.addEventListener('click', showSettingsModal);
        tabButtons.forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.view)));
        inventoryFilters.addEventListener('click', (e) => { 
            const button = e.target.closest('.filter-btn');
            if (!button) return;
            inventoryFilter = button.dataset.filter;
            document.querySelectorAll('#inventory-filters .filter-btn').forEach(btn => btn.classList.remove('active', 'bg-white', 'text-indigo-600', 'shadow-sm'));
            button.classList.add('active', 'bg-white', 'text-indigo-600', 'shadow-sm');
            currentPage = 1; // 🚀 Resetar para primeira página ao filtrar
            renderProducts();
        });
        sortOrderSelect.addEventListener('change', (e) => { 
            inventorySortOrder = e.target.value;
            currentPage = 1; // 🚀 Resetar para primeira página ao ordenar
            renderProducts();
        });
        
        micBtn.addEventListener('click', () => {
            if (!recognition) return;
            if (micBtn.classList.contains('is-listening')) {
                recognition.stop();
            } else {
                recognition.start();
                micBtn.classList.add('is-listening');
                document.getElementById('mic-status').textContent = 'Ouvindo...';
            }
        });

        productList.addEventListener('click', async (e) => {
            const button = e.target.closest('button');
            if (button) {
                currentProductId = button.dataset.id;

                // 🔐 Verificar permissões para operações
                if (button.classList.contains('edit-btn')) {
                    if (!hasPermission('update')) {
                        showToast('🔒 Você não tem permissão para editar produtos', true);
                        return;
                    }
                    showEditModal();
                }
                
                if (button.classList.contains('history-btn')) showHistoryModal();
                
                if (button.classList.contains('delete-btn')) {
                    if (!hasPermission('delete')) {
                        showToast('🔒 Você não tem permissão para excluir produtos', true);
                        return;
                    }
                    showConfirmationModal(
                        'Excluir Produto',
                        'Tem certeza que deseja excluir este produto? Esta ação não pode ser desfeita.',
                        async () => {
                            try {
                                await deleteDoc(doc(productsCollectionRef, currentProductId));
                                selectedProductIds.delete(currentProductId);
                                showToast("Produto excluído.");
                            } catch (error) { showToast("Falha ao excluir produto.", true); }
                        }
                    );
                }
            }
        });
        
        productList.addEventListener('change', (e) => {
            if (e.target.classList.contains('product-checkbox')) {
                const productId = e.target.dataset.id;
                if (e.target.checked) {
                    selectedProductIds.add(productId);
                } else {
                    selectedProductIds.delete(productId);
                }
                updateSelectionActionButtonsState();
            }
        });

        selectAllProductsCheckbox.addEventListener('change', (e) => {
            const isChecked = e.target.checked;
            const displayedCheckboxes = document.querySelectorAll('.product-checkbox');
            
            displayedCheckboxes.forEach(checkbox => {
                const productId = checkbox.dataset.id;
                checkbox.checked = isChecked;
                if (isChecked) {
                    selectedProductIds.add(productId);
                } else {
                    selectedProductIds.delete(productId);
                }
            });
            updateSelectionActionButtonsState();
        });

        deleteSelectedBtn.addEventListener('click', () => {
            if (selectedProductIds.size === 0) return;

            showConfirmationModal(
                `Excluir ${selectedProductIds.size} Produtos`,
                'Tem certeza que deseja excluir os produtos selecionados? Esta ação não pode ser desfeita.',
                async () => {
                    const batch = writeBatch(db);
                    selectedProductIds.forEach(id => {
                        batch.delete(doc(productsCollectionRef, id));
                    });
                    try {
                        await batch.commit();
                        showToast(`${selectedProductIds.size} produtos foram excluídos.`);
                        selectedProductIds.clear();
                        renderProducts();
                    } catch (error) {
                        console.error("Erro ao excluir produtos em lote: ", error);
                        showToast("Falha ao excluir produtos.", true);
                    }
                }
            );
        });
        
        initiateRequisitionBtn.addEventListener('click', () => {
            if (selectedProductIds.size === 0) {
                showToast("Selecione produtos para iniciar uma requisição.", true);
                return;
            }
            const idsToRequisition = new Set(selectedProductIds);
            switchView('requisitions-view');
            showRequisitionModal(idsToRequisition);
        });

        addForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            // 🔐 Verificar permissão
            if (!hasPermission('create')) {
                showToast('🔒 Você não tem permissão para adicionar produtos', true);
                return;
            }

            const newProduct = {
                code: document.getElementById('product-code').value.trim(),
                codeRM: document.getElementById('product-code-rm').value.trim(),
                name: document.getElementById('product-name').value.trim(),
                unit: document.getElementById('product-unit').value,
                group: document.getElementById('product-group').value,
                quantity: parseInt(document.getElementById('product-quantity').value),
                minQuantity: parseInt(document.getElementById('product-min-quantity').value),
                location: document.getElementById('product-location').value.trim(),
                observation: '',
                createdAt: serverTimestamp(),
                updatedAt: serverTimestamp(),
                createdBy: currentUser?.displayName || currentUser?.uid || 'Anônimo'
            };

            // 🔒 VALIDAÇÕES CRÍTICAS
            // 1. Campos obrigatórios
            if (!newProduct.name) {
                showToast('❌ Nome do Produto é obrigatório.', true);
                return;
            }

            // 2. Validar quantidades não negativas
            if (newProduct.quantity < 0 || newProduct.minQuantity < 0) {
                showToast('❌ Quantidades não podem ser negativas.', true);
                return;
            }

            // 3. Validar se quantidades são números válidos
            if (isNaN(newProduct.quantity) || isNaN(newProduct.minQuantity)) {
                showToast('❌ Por favor, insira quantidades válidas.', true);
                return;
            }

            // 4. Verificar código RM duplicado (apenas se preenchido)
            if (newProduct.codeRM) {
                const duplicateRM = products.find(p => 
                    p.codeRM && p.codeRM.toLowerCase() === newProduct.codeRM.toLowerCase()
                );
                if (duplicateRM) {
                    showToast(`❌ Código RM "${newProduct.codeRM}" já existe no produto: ${duplicateRM.name}`, true);
                    return;
                }
            }

            // 5. Verificar código duplicado (se preenchido)
            if (newProduct.code) {
                const duplicateCode = products.find(p => 
                    p.code && p.code.toLowerCase() === newProduct.code.toLowerCase()
                );
                if (duplicateCode) {
                    showToast(`❌ Código "${newProduct.code}" já existe no produto: ${duplicateCode.name}`, true);
                    return;
                }
            }

            try {
                const docRef = await addDoc(productsCollectionRef, newProduct);
                await addHistoryEntry(docRef.id, 'Criação', newProduct.quantity, newProduct.quantity, {}, newProduct);
                showToast("✅ Produto adicionado com sucesso!");
                addForm.reset();
                switchView('inventory-view');
            } catch (error) {
                console.error("Erro ao adicionar produto: ", error);
                const errorMsg = error.code === 'permission-denied' 
                    ? '🔒 Sem permissão para adicionar produtos. Verifique as configurações do Firebase.'
                    : '❌ Falha ao adicionar produto. Tente novamente.';
                showToast(errorMsg, true);
            }
        });
        
        createRequisitionBtn.addEventListener('click', () => {
            selectedProductIds.clear();
            showRequisitionModal();
        });
        
        requisitionsList.addEventListener('click', e => {
            const button = e.target.closest('.view-requisition-btn');
            if (button) showViewRequisitionModal([button.dataset.id]);
        });
        
        document.body.addEventListener('submit', async (e) => {
            if (e.target.id === 'adjust-form') {
                e.preventDefault();
                const type = e.target.querySelector('input[name="adjust-type"]:checked').value;
                const quantityChange = parseInt(document.getElementById('quantity-change').value);

                if (isNaN(quantityChange) || quantityChange <= 0) {
                    showToast("Por favor, insira uma quantidade válida.", true);
                    return;
                }

                const productRef = doc(productsCollectionRef, currentProductId);

                try {
                    let productDataForHistory = null;
                    let newQuantity = 0;
                    
                    await runTransaction(db, async (transaction) => {
                        const productDoc = await transaction.get(productRef);
                        if (!productDoc.exists()) {
                            throw new Error("Produto não encontrado.");
                        }
                        
                        productDataForHistory = productDoc.data();
                        const currentQuantity = productDataForHistory.quantity;
                        
                        if (type === 'entrada') {
                            newQuantity = currentQuantity + quantityChange;
                        } else {
                            if (currentQuantity < quantityChange) {
                                throw new Error("Estoque insuficiente para este ajuste de saída.");
                            }
                            newQuantity = currentQuantity - quantityChange;
                        }
                        
                        transaction.update(productRef, { quantity: newQuantity });
                    });

                    await addHistoryEntry(
                        currentProductId, 
                        type === 'entrada' ? 'Ajuste Entrada' : 'Ajuste Saída',
                        quantityChange,
                        newQuantity,
                        {},
                        productDataForHistory
                    );

                    showToast("Estoque ajustado com sucesso!");
                    closeModal('generic-modal');
                } catch (error) {
                    console.error("Erro ao ajustar estoque: ", error);
                    showToast(`Falha ao ajustar estoque: ${error.message}`, true);
                }
            }
            
            if (e.target.id === 'requisition-form') {
                e.preventDefault();
                const button = e.target.querySelector('button[type="submit"]');
                button.disabled = true;
                button.innerHTML = `<div class="spinner-small"></div>`;

                const itemsMap = new Map();
                let hasInvalidItem = false;

                document.querySelectorAll('#req-items-container .flex').forEach(row => {
                    const productId = row.querySelector('.req-item-product').value;
                    const quantity = parseInt(row.querySelector('.req-item-quantity').value);
                    if (!productId || isNaN(quantity) || quantity <= 0) {
                        hasInvalidItem = true;
                    }
                    if (productId) {
                        const existing = itemsMap.get(productId) || { quantity: 0 };
                        itemsMap.set(productId, { quantity: existing.quantity + quantity });
                    }
                });

                if (hasInvalidItem || itemsMap.size === 0) {
                    showToast("Adicione pelo menos um item com quantidade e produto válidos.", true);
                    button.disabled = false;
                    button.textContent = 'Salvar e Baixar Estoque';
                    return;
                }
                
                const newReqData = {
                    number: document.getElementById('req-number').value,
                    requester: document.getElementById('req-requester').value.trim(),
                    teamLeader: document.getElementById('req-team-leader').value.trim(),
                    applicationLocation: document.getElementById('req-application-location').value,
                    obra: document.getElementById('req-obra').value,
                    items: [], 
                    date: serverTimestamp()
                };

                try {
                    const historyEntries = [];
                    await runTransaction(db, async (transaction) => {
                        const productRefs = Array.from(itemsMap.keys()).map(id => doc(productsCollectionRef, id));
                        const productDocs = await Promise.all(productRefs.map(ref => transaction.get(ref)));
                        
                        for (let i = 0; i < productDocs.length; i++) {
                            const productDoc = productDocs[i];
                            const productId = productRefs[i].id;
                            const requestedQuantity = itemsMap.get(productId).quantity;

                            if (!productDoc.exists()) throw new Error(`Produto não encontrado no banco de dados.`);
                            
                            const productData = productDoc.data();
                            if (productData.quantity < requestedQuantity) throw new Error(`Estoque insuficiente para ${productData.name}.`);
                            
                            const newQuantity = productData.quantity - requestedQuantity;
                            transaction.update(productDoc.ref, { quantity: newQuantity });
                            
                            newReqData.items.push({
                                productId,
                                productName: productData.name,
                                productCode: productData.code,
                                productCodeRM: productData.codeRM,
                                quantity: requestedQuantity
                            });

                            historyEntries.push({
                                productId,
                                productData,
                                type: 'Saída por Requisição',
                                quantity: requestedQuantity,
                                newTotal: newQuantity,
                                details: {
                                    withdrawnBy: newReqData.requester,
                                    teamLeader: newReqData.teamLeader,
                                    applicationLocation: newReqData.applicationLocation,
                                    obra: newReqData.obra,
                                    details: `Requisição Nº ${newReqData.number}`
                                }
                            });
                        }
                        transaction.set(doc(requisitionsCollectionRef), newReqData);
                    });

                    for(const entry of historyEntries) {
                        const { productId, type, quantity, newTotal, details, productData } = entry;
                         await addHistoryEntry(productId, type, quantity, newTotal, details, productData);
                    }
                    
                    selectedProductIds.clear();
                    showToast("Requisição criada e estoque atualizado!");
                    closeModal('generic-modal');

                } catch (error) {
                    console.error("Erro na transação da requisição: ", error);
                    showToast(`Erro: ${error.message}`, true);
                } finally {
                    button.disabled = false;
                    button.textContent = 'Salvar e Baixar Estoque';
                }
            }

            if (e.target.id === 'edit-product-form') {
                e.preventDefault();
                const id = document.getElementById('edit-product-id').value;
                const productRef = doc(productsCollectionRef, id);
                const newQuantity = parseInt(document.getElementById('edit-product-quantity').value);
                const updatedData = {
                    codeRM: document.getElementById('edit-product-code-rm').value.trim(),
                    name: document.getElementById('edit-product-name').value.trim(),
                    unit: document.getElementById('edit-product-unit').value,
                    group: document.getElementById('edit-product-group').value,
                    quantity: newQuantity,
                    minQuantity: parseInt(document.getElementById('edit-product-min-quantity').value),
                    location: document.getElementById('edit-product-location').value.trim(),
                    observation: document.getElementById('edit-product-observation').value.trim(),
                    updatedAt: serverTimestamp(), // 🔍 Auditoria: timestamp de atualização
                    updatedBy: currentUser?.displayName || currentUser?.uid || 'Anônimo' // 🔍 Auditoria: quem atualizou
                };
                try {
                    await updateDoc(productRef, updatedData);
                    const product = products.find(p => p.id === id);
                    await addHistoryEntry(id, 'Edição', 0, newQuantity, { details: 'Dados do produto alterados.' }, { ...product, ...updatedData });
                    showToast("Produto atualizado com sucesso!");
                    closeModal('generic-modal');
                } catch (error) {
                    console.error("Erro ao editar produto: ", error);
                    showToast("Falha ao editar produto.", true);
                }
            }
            if (e.target.id === 'settings-form') {
                 e.preventDefault();
                const newName = document.getElementById('settings-app-name').value.trim();
                const logoPreview = document.getElementById('logo-preview');
                const newLogoUrl = logoPreview.src.startsWith('data:image') ? logoPreview.src : (logoPreview.dataset.removed ? null : appSettings.logoUrl);

                if (!newName) {
                    showToast("O nome do app não pode ficar em branco.", true);
                    return;
                }
                const updatedSettings = {
                    appName: newName,
                    logoUrl: newLogoUrl
                };
                try {
                    await setDoc(settingsDocRef, updatedSettings, { merge: true });
                    showToast("Configurações salvas.");
                    closeModal('generic-modal');
                } catch (error) {
                    console.error("Erro ao salvar configurações: ", error);
                    showToast("Falha ao salvar configurações.", true);
                }
            }
            if (e.target.id === 'add-location-form') {
                e.preventDefault();
                const name = document.getElementById('location-name').value.trim();
                const code = document.getElementById('location-code').value.trim();
                const aliasesInput = document.getElementById('location-aliases').value.trim();
                const aliases = aliasesInput ? aliasesInput.split(',').map(a => a.trim()).filter(Boolean) : [];

                if (!name) { showToast("O nome do local é obrigatório.", true); return; }
                try {
                    await addDoc(locationsCollectionRef, { name, code, aliases });
                    showToast("Local adicionado com sucesso!");
                    addLocationForm.reset();
                } catch (error) {
                    console.error("Erro ao adicionar local:", error);
                    showToast("Falha ao adicionar local.", true);
                }
            }
            if (e.target.id === 'edit-location-form') {
                e.preventDefault();
                const id = document.getElementById('edit-location-id').value;
                const locationRef = doc(locationsCollectionRef, id);
                const aliasesInput = document.getElementById('edit-location-aliases').value.trim();
                const aliases = aliasesInput ? aliasesInput.split(',').map(a => a.trim()).filter(Boolean) : [];
                const updatedData = {
                    name: document.getElementById('edit-location-name').value.trim(),
                    code: document.getElementById('edit-location-code').value.trim(),
                    aliases: aliases
                };
                try {
                    await updateDoc(locationRef, updatedData);
                    showToast("Local atualizado com sucesso!");
                    closeModal('generic-modal');
                } catch (error) {
                    console.error("Erro ao editar local: ", error);
                    showToast("Falha ao editar local.", true);
                }
            }
            if (e.target.id === 'entry-form') {
                e.preventDefault();
                const productId = document.getElementById('entry-product-select').value;
                const quantity = parseInt(document.getElementById('entry-quantity').value);
                const nfNumber = document.getElementById('entry-nf').value.trim();
                const supplier = document.getElementById('entry-supplier').value.trim();
                const observation = document.getElementById('entry-observation').value.trim();
                const receivedBy = currentUser?.displayName || currentUser?.uid || 'Sistema';

                if (!productId || !nfNumber || isNaN(quantity) || quantity <= 0) {
                    showToast("Por favor, preencha todos os campos obrigatórios corretamente.", true);
                    return;
                }
                
                const button = e.target.querySelector('button[type="submit"]');
                button.disabled = true;
                button.innerHTML = `<div class="spinner-small"></div>`;

                const productRef = doc(productsCollectionRef, productId);
                try {
                    let productDataForHistory = null;
                    await runTransaction(db, async (transaction) => {
                        const productDoc = await transaction.get(productRef);
                        if (!productDoc.exists()) {
                            throw new Error("Produto não encontrado.");
                        }
                        productDataForHistory = productDoc.data();
                        const currentQuantity = productDataForHistory.quantity;
                        const newQuantity = currentQuantity + quantity;
                        transaction.update(productRef, { quantity: newQuantity });
                        
                        const historyRef = doc(historyCollectionRef);
                        transaction.set(historyRef, {
                            productId,
                            productCode: productDataForHistory.code,
                            productCodeRM: productDataForHistory.codeRM,
                            productName: productDataForHistory.name,
                            type: 'Entrada por NF',
                            quantity: quantity,
                            newTotal: newQuantity,
                            receivedBy,
                            supplier,
                            nfNumber,
                            observation,
                            details: observation ? `Entrada da NF ${nfNumber} - ${observation}` : `Entrada da NF ${nfNumber}`,
                            rmProcessed: false,
                            date: serverTimestamp()
                        });
                    });

                    showToast(`Entrada de ${quantity} unidade(s) de ${productDataForHistory.name} registrada com sucesso! Novo estoque: ${productDataForHistory.quantity + quantity}.`);
                    e.target.reset();
                    document.getElementById('entry-product-select').value = '';

                } catch (error) {
                    console.error("Erro ao registrar entrada: ", error);
                    showToast(`Falha ao registrar entrada: ${error.message}`, true);
                } finally {
                    button.disabled = false;
                    button.textContent = 'Confirmar Entrada';
                }
            }
            if (e.target.id === 'exit-form') {
                e.preventDefault();
                const productId = document.getElementById('exit-product-select').value;
                const quantity = parseInt(document.getElementById('exit-quantity').value);
                const withdrawnBy = document.getElementById('exit-withdrawn-by').value.trim();
                const obra = document.getElementById('exit-obra').value;
                const applicationLocation = document.getElementById('exit-application-location').value.trim();

                if (!productId || !withdrawnBy || !applicationLocation || isNaN(quantity) || quantity <= 0) {
                    showToast("Por favor, preencha todos os campos corretamente.", true);
                    return;
                }
                
                const button = e.target.querySelector('button[type="submit"]');
                button.disabled = true;
                button.innerHTML = `<div class="spinner-small"></div>`;

                const productRef = doc(productsCollectionRef, productId);
                try {
                    let productDataForHistory = null;
                    await runTransaction(db, async (transaction) => {
                        const productDoc = await transaction.get(productRef);
                        if (!productDoc.exists()) {
                            throw new Error("Produto não encontrado.");
                        }
                        productDataForHistory = productDoc.data();
                        const currentQuantity = productDataForHistory.quantity;
                        if (currentQuantity < quantity) {
                            throw new Error(`Estoque insuficiente. Disponível: ${currentQuantity}`);
                        }
                        const newQuantity = currentQuantity - quantity;
                        transaction.update(productRef, { quantity: newQuantity });
                        
                        const historyRef = doc(historyCollectionRef);
                        transaction.set(historyRef, {
                            productId,
                            productCode: productDataForHistory.code,
                            productCodeRM: productDataForHistory.codeRM,
                            productName: productDataForHistory.name,
                            type: 'Saída',
                            quantity: quantity,
                            newTotal: newQuantity,
                            withdrawnBy,
                            obra,
                            applicationLocation,
                            teamLeader: null,
                            details: 'Saída manual',
                            rmProcessed: false,
                            date: serverTimestamp()
                        });
                    });

                    showToast("Saída registrada com sucesso!");
                    e.target.reset();
                    document.getElementById('ai-exit-prompt').value = '';
                    document.getElementById('exit-product-select').value = '';

                } catch (error) {
                    console.error("Erro ao registrar saída: ", error);
                    showToast(`Falha ao registrar saída: ${error.message}`, true);
                } finally {
                    button.disabled = false;
                    button.textContent = 'Confirmar Saída';
                }
            }
        });

        document.body.addEventListener('click', e => {
            if (e.target.closest('.close-modal-btn')) closeModal('generic-modal');
            if (e.target.id === 'go-to-add-product') switchView('add-product-view');
            if (e.target.id === 'upload-logo-btn') document.getElementById('logo-upload-input').click();
            if (e.target.id === 'remove-logo-btn') {
                const preview = document.getElementById('logo-preview');
                preview.src = 'https://placehold.co/200x60/e2e8f0/475569?text=Sem+Logo';
                preview.dataset.removed = "true";
            }
            if (e.target.classList.contains('copy-btn')) {
                copyToClipboard(e.target.dataset.copyText);
            }
        });

        document.body.addEventListener('change', e => {
            if (e.target.id === 'logo-upload-input') {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (event) => {
                        const preview = document.getElementById('logo-preview');
                        preview.src = event.target.result;
                        preview.dataset.removed = "false";
                    };
                    reader.readAsDataURL(file);
                }
            }
            if (e.target.classList.contains('rm-status-checkbox')) {
                const historyId = e.target.dataset.id;
                const isChecked = e.target.checked;
                const historyRef = doc(historyCollectionRef, historyId);
                updateDoc(historyRef, { rmProcessed: isChecked })
                    .then(() => {
                        showToast(`Status atualizado.`);
                    })
                    .catch(error => {
                        console.error("Erro ao atualizar status: ", error);
                        showToast("Falha ao atualizar status.", true);
                        e.target.checked = !isChecked;
                    });
            }
        });


        locationsList.addEventListener('click', async (e) => {
            const button = e.target.closest('button');
            if (!button) return;
            currentLocationId = button.dataset.id;
            
            if (button.classList.contains('edit-location-btn')) {
                showEditLocationModal();
            }
            
            if (button.classList.contains('delete-location-btn')) {
                showConfirmationModal(
                    'Excluir Local',
                    'Tem certeza que deseja excluir este local? Esta ação não pode ser desfeita.',
                    async () => {
                        try {
                            await deleteDoc(doc(locationsCollectionRef, currentLocationId));
                            showToast("Local excluído.");
                        } catch (error) { 
                            console.error("Erro ao excluir local:", error);
                            showToast("Falha ao excluir local.", true); 
                        }
                    }
                );
            }
        });

        importBtn.addEventListener('click', () => csvFileInput.click());
        csvFileInput.addEventListener('change', event => {
            const file = event.target.files[0]; if (!file) return;
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const lines = e.target.result.split(/\r\n|\n/);
                    let importedCount = 0;
                    if (lines.length < 2) { showToast('Arquivo CSV vazio ou mal formatado.', true); return; }
                    
                    const delimiter = lines[0].includes(';') ? ';' : ',';
                    const batch = writeBatch(db);
                    let nextCode = parseInt(generateNextProductCode());

                    for (let i = 1; i < lines.length; i++) {
                        if (lines[i].trim() === '') continue;
                        const columns = lines[i].split(delimiter);
                        if (columns.length < 6) continue;

                        const [codeRM, name, unitAbbr, quantity, location, minQuantity, group] = columns.map(s => s.trim().replace(/"/g, ''));
                        const unit = unitMap[unitAbbr.toLowerCase()] || unitAbbr;

                        if (codeRM && name) {
                            const newProd = {
                                code: (nextCode++).toString(),
                                codeRM, name, unit, location,
                                group: group || 'Outros',
                                quantity: parseInt(quantity) || 0,
                                minQuantity: parseInt(minQuantity) || 0,
                                observation: ''
                            };
                            const docRef = doc(productsCollectionRef);
                            batch.set(docRef, newProd);
                            importedCount++;
                        }
                    }
                    
                    await batch.commit();

                    if (importedCount > 0) {
                        showToast(`${importedCount} produtos foram importados com sucesso!`);
                    } else {
                        showToast('Nenhum produto novo foi importado.', true);
                    }
                } catch (error) {
                    showToast('Ocorreu um erro ao ler o arquivo.', true);
                    console.error(error);
                } finally {
                    csvFileInput.value = '';
                }
            };
            reader.readAsText(file, 'UTF-8');
        });

        importLocationsBtn.addEventListener('click', () => csvLocationsFileInput.click());
        csvLocationsFileInput.addEventListener('change', event => {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const lines = e.target.result.split(/\r\n|\n/);
                    let importedCount = 0;
                    if (lines.length < 2) {
                        showToast('Arquivo CSV vazio ou mal formatado.', true);
                        return;
                    }
                    
                    const delimiter = lines[0].includes(';') ? ';' : ',';
                    const batch = writeBatch(db);

                    for (let i = 1; i < lines.length; i++) {
                        if (lines[i].trim() === '') continue;
                        const columns = lines[i].split(delimiter);
                        if (columns.length < 1) continue;

                        const [name, code] = columns.map(s => s.trim().replace(/"/g, ''));

                        if (name) {
                            const newLoc = {
                                name: name,
                                code: code || ''
                            };
                            const docRef = doc(locationsCollectionRef);
                            batch.set(docRef, newLoc);
                            importedCount++;
                        }
                    }
                    
                    await batch.commit();

                    if (importedCount > 0) {
                        showToast(`${importedCount} locais foram importados com sucesso!`);
                    } else {
                        showToast('Nenhum local novo foi importado.', true);
                    }
                } catch (error) {
                    showToast('Ocorreu um erro ao ler o arquivo.', true);
                    console.error(error);
                } finally {
                    csvLocationsFileInput.value = '';
                }
            };
            reader.readAsText(file, 'UTF-8');
        });

        exportBtn.addEventListener('click', () => {
            if (products.length === 0) { showToast("Não há produtos para exportar.", true); return; }
            let csvContent = "Código RM;Nome;Unidade;Quantidade;Localização;Qtd. Mínima;Grupo\n";
            products.forEach(p => { 
                const row = [
                    p.codeRM || '',
                    p.name || '',
                    p.unit || '',
                    p.quantity || 0,
                    p.location || '',
                    p.minQuantity || 0,
                    p.group || ''
                ].map(field => `"${String(field).replace(/"/g, '""')}"`).join(';');
                csvContent += row + '\n';
            });
            const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", "estoque_exportado.csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });

        printSelectedReqsBtn.addEventListener('click', () => {
            const selectedIds = Array.from(document.querySelectorAll('.req-checkbox:checked')).map(cb => cb.dataset.id);
            if (selectedIds.length === 0) {
                showToast("Selecione pelo menos uma requisição para imprimir.", true);
                return;
            }
            showViewRequisitionModal(selectedIds);
        });
        document.getElementById('select-all-reqs').addEventListener('change', (e) => {
            document.querySelectorAll('.req-checkbox').forEach(checkbox => {
                checkbox.checked = e.target.checked;
            });
        });
        
        generateExitReportBtn.addEventListener('click', () => {
            const exitHistory = history.filter(h => h.type === 'Saída' || h.type === 'Saída por Requisição');
            if (exitHistory.length === 0) {
                showToast("Nenhum registro de saída para exportar.", true);
                return;
            }

            let csvContent = "Data;Produto;Código RM;Cód. Interno;Grupo;Quantidade;Retirado por;Líder da Equipe;Obra;Local de Aplicação;Código do Local;Detalhes\n";
            exitHistory.sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0)).forEach(h => {
                const date = h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : 'N/A';
                const location = locations.find(loc => loc.name === h.applicationLocation);
                const locationCode = location ? location.code : 'N/A';
                const product = products.find(p => p.id === h.productId);
                const row = [
                    date,
                    h.productName || '',
                    h.productCodeRM || '',
                    h.productCode || '',
                    product?.group || '',
                    h.quantity || 0,
                    h.withdrawnBy || '',
                    h.teamLeader || '',
                    h.obra || '',
                    h.applicationLocation || '',
                    locationCode,
                    h.details || ''
                ].map(field => `"${String(field).replace(/"/g, '""')}"`).join(';');
                csvContent += row + '\n';
            });

            const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement("a");
            const url = URL.createObjectURL(blob);
            link.setAttribute("href", url);
            link.setAttribute("download", "relatorio_saidas.csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });

        generateConsumptionReportBtn.addEventListener('click', () => {
            const exitHistory = history.filter(h => h.type === 'Saída' || h.type === 'Saída por Requisição');
            if (exitHistory.length === 0) {
                showToast("Nenhum consumo registrado para gerar relatório.", true);
                return;
            }

            const consumptionByLocation = exitHistory.reduce((acc, curr) => {
                const location = curr.applicationLocation || 'Não especificado';
                if (!acc[location]) {
                    acc[location] = {};
                }
                const productIdentifier = curr.productCodeRM || curr.productCode;
                if (!acc[location][productIdentifier]) {
                    const product = products.find(p => p.id === curr.productId);
                    acc[location][productIdentifier] = {
                        name: curr.productName,
                        codeRM: curr.productCodeRM,
                        group: product?.group || 'N/A',
                        unit: product?.unit || 'N/A',
                        totalQuantity: 0
                    };
                }
                acc[location][productIdentifier].totalQuantity += curr.quantity;
                return acc;
            }, {});

            let modalContent = `
                <div class="flex justify-between items-start mb-6">
                    <h2 class="text-2xl font-bold">Relatório de Consumo por Local</h2>
                    <button type="button" class="close-modal-btn p-2 -mt-2 -mr-2 text-slate-400 hover:text-slate-700 text-3xl">&times;</button>
                </div>
                <div id="consumption-report-content" class="max-h-96 overflow-y-auto space-y-6 pr-2">
            `;

            for (const location in consumptionByLocation) {
                modalContent += `
                    <div class="mb-4">
                        <h3 class="text-lg font-bold text-indigo-700 border-b pb-2 mb-2">${location}</h3>
                        <table class="w-full text-left text-sm">
                            <thead class="bg-slate-50">
                                <tr>
                                    <th class="p-2">Produto</th>
                                    <th class="p-2">Cód. RM</th>
                                    <th class="p-2">Grupo</th>
                                    <th class="p-2 text-center">Qtd. Consumida</th>
                                    <th class="p-2">Unidade</th>
                                </tr>
                            </thead>
                            <tbody>
                `;
                for (const product in consumptionByLocation[location]) {
                    const data = consumptionByLocation[location][product];
                    modalContent += `
                        <tr class="border-b">
                            <td class="p-2">${data.name}</td>
                            <td class="p-2">${data.codeRM}</td>
                            <td class="p-2">${data.group}</td>
                            <td class="p-2 text-center font-semibold">${data.totalQuantity}</td>
                            <td class="p-2">${data.unit}</td>
                        </tr>
                    `;
                }
                modalContent += `</tbody></table></div>`;
            }
            modalContent += `</div>
                <div class="mt-8 flex justify-end gap-4">
                    <button type="button" id="export-consumption-csv" class="px-6 py-2 bg-teal-600 text-white rounded-lg font-semibold hover:bg-teal-700 transition">Exportar (CSV)</button>
                </div>
            `;
            
            document.getElementById('generic-modal').innerHTML = modalContent;
            document.getElementById('generic-modal').classList.replace('max-w-md', 'max-w-4xl');
            openModal('generic-modal');

            document.getElementById('export-consumption-csv').addEventListener('click', () => {
                let csvContent = "Local de Aplicação;Produto;Código RM;Grupo;Unidade;Quantidade Consumida\n";
                for (const location in consumptionByLocation) {
                    for (const product in consumptionByLocation[location]) {
                        const data = consumptionByLocation[location][product];
                        const row = [location, data.name, data.codeRM, data.group, data.unit, data.totalQuantity]
                            .map(field => `"${String(field).replace(/"/g, '""')}"`).join(';');
                        csvContent += row + '\n';
                    }
                }
                const blob = new Blob([`\uFEFF${csvContent}`], { type: 'text/csv;charset=utf-8;' });
                const link = document.createElement("a");
                const url = URL.createObjectURL(blob);
                link.setAttribute("href", url);
                link.setAttribute("download", "relatorio_consumo_por_local.csv");
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });
        });

        generateKpiReportBtn.addEventListener('click', () => {
            const totalProducts = products.length;
            const totalUnits = products.reduce((sum, p) => sum + (p.quantity || 0), 0);
            const lowStockItems = products.filter(p => p.quantity <= p.minQuantity);
            const lowStockCount = lowStockItems.length;
            const pendingReqs = requisitions.filter(r => r.status === 'Pendente' || r.status === 'pending').length;

            const timestampToMillis = (entry) => {
                if (!entry?.date) return 0;
                if (typeof entry.date.seconds === 'number') return entry.date.seconds * 1000;
                if (typeof entry.date.toMillis === 'function') return entry.date.toMillis();
                return 0;
            };

            const getTopExitedItems = (days) => {
                const cutoff = Date.now() - (days * 24 * 60 * 60 * 1000);
                const grouped = new Map();

                history
                    .filter(h => (h.type === 'Saída' || h.type === 'Saída por Requisição') && timestampToMillis(h) >= cutoff)
                    .forEach(h => {
                        const key = h.productId || h.productCode || h.productCodeRM || h.productName;
                        const previous = grouped.get(key) || {
                            name: h.productName || 'Produto sem nome',
                            code: h.productCodeRM || h.productCode || 'N/A',
                            qty: 0
                        };
                        previous.qty += Math.abs(Number(h.quantity) || 0);
                        grouped.set(key, previous);
                    });

                return [...grouped.values()].sort((a, b) => b.qty - a.qty).slice(0, 10);
            };

            const getEntriesCount = (days) => {
                const cutoff = Date.now() - (days * 24 * 60 * 60 * 1000);
                return history.filter(h => (h.type === 'Entrada' || h.type === 'Ajuste Entrada') && timestampToMillis(h) >= cutoff)
                    .reduce((s, h) => s + Math.abs(Number(h.quantity) || 0), 0);
            };
            const getExitsCount = (days) => {
                const cutoff = Date.now() - (days * 24 * 60 * 60 * 1000);
                return history.filter(h => (h.type === 'Saída' || h.type === 'Saída por Requisição') && timestampToMillis(h) >= cutoff)
                    .reduce((s, h) => s + Math.abs(Number(h.quantity) || 0), 0);
            };

            const top3Days = getTopExitedItems(3);
            const top7Days = getTopExitedItems(7);
            const top30Days = getTopExitedItems(30);

            const entries7d = getEntriesCount(7);
            const exits7d = getExitsCount(7);
            const entries30d = getEntriesCount(30);
            const exits30d = getExitsCount(30);

            // Grupos de produtos
            const groupMap = new Map();
            products.forEach(p => {
                const g = p.group || 'Sem Grupo';
                const prev = groupMap.get(g) || { count: 0, units: 0, lowStock: 0 };
                prev.count++;
                prev.units += (p.quantity || 0);
                if (p.quantity <= p.minQuantity) prev.lowStock++;
                groupMap.set(g, prev);
            });
            const groupData = [...groupMap.entries()].sort((a, b) => b[1].units - a[1].units);

            const esc = v => String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            const dataGerada = new Date().toLocaleString('pt-BR');
            const timestamp = new Date().toISOString().slice(0, 10);

            // Estilo base das células
            const hd = 'background:#005BBF;color:#FFFFFF;font-weight:bold;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;';
            const cd = 'border:1px solid #D0D0D0;padding:6px 8px;font-size:11px;font-family:Calibri,sans-serif;';
            const secTitle = 'font-size:14px;font-weight:bold;color:#005BBF;padding:10px 8px 4px 8px;font-family:Calibri,sans-serif;';

            const renderTopTable = (items, periodLabel) => {
                if (!items.length) return `<tr><td colspan="4" style="${cd}color:#999;">Sem registros no período</td></tr>`;
                return items.map((item, idx) => {
                    const bg = idx % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
                    const medal = idx === 0 ? '🥇' : idx === 1 ? '🥈' : idx === 2 ? '🥉' : `${idx + 1}º`;
                    return `<tr>
                        <td style="${cd}background:${bg};text-align:center;font-weight:bold;">${medal}</td>
                        <td style="${cd}background:${bg};">${esc(item.name)}</td>
                        <td style="${cd}background:${bg};text-align:center;">${esc(item.code)}</td>
                        <td style="${cd}background:${bg};text-align:center;font-weight:bold;color:#B71C1C;">${item.qty}</td>
                    </tr>`;
                }).join('');
            };

            // Low stock items table
            const lowStockSorted = [...lowStockItems].sort((a, b) => (a.quantity - a.minQuantity) - (b.quantity - b.minQuantity));
            const lowStockRows = lowStockSorted.slice(0, 15).map((p, idx) => {
                const bg = idx % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
                const deficit = (p.minQuantity || 0) - (p.quantity || 0);
                const deficitColor = deficit > 0 ? '#B71C1C' : '#E65100';
                return `<tr>
                    <td style="${cd}background:${bg};">${esc(p.name)}</td>
                    <td style="${cd}background:${bg};text-align:center;">${esc(p.codeRM || p.code || '-')}</td>
                    <td style="${cd}background:${bg};text-align:center;">${esc(p.group || '-')}</td>
                    <td style="${cd}background:${bg};text-align:center;">${p.quantity ?? 0}</td>
                    <td style="${cd}background:${bg};text-align:center;">${p.minQuantity ?? 0}</td>
                    <td style="${cd}background:${bg};text-align:center;font-weight:bold;color:${deficitColor};">${deficit > 0 ? `-${deficit}` : '0'}</td>
                </tr>`;
            }).join('');

            // Group summary rows
            const groupRows = groupData.map(([ g, d ], idx) => {
                const bg = idx % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
                return `<tr>
                    <td style="${cd}background:${bg};font-weight:bold;">${esc(g)}</td>
                    <td style="${cd}background:${bg};text-align:center;">${d.count}</td>
                    <td style="${cd}background:${bg};text-align:center;">${d.units.toLocaleString('pt-BR')}</td>
                    <td style="${cd}background:${bg};text-align:center;color:${d.lowStock > 0 ? '#B71C1C' : '#333'};font-weight:${d.lowStock > 0 ? 'bold' : 'normal'};">${d.lowStock}</td>
                </tr>`;
            }).join('');

            const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8">
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>KPIs Estoque</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
</head><body>
<table>
    <!-- CABEÇALHO -->
    <tr><td colspan="6" style="font-size:20px;font-weight:bold;color:#005BBF;padding:14px 8px 2px 8px;font-family:Calibri,sans-serif;">RELATÓRIO DE KPIs — ESTOQUE UHE ESTRELA</td></tr>
    <tr><td colspan="6" style="font-size:11px;color:#555;padding:2px 8px 10px 8px;font-family:Calibri,sans-serif;">Gerado em: ${esc(dataGerada)}</td></tr>

    <!-- SEÇÃO 1: INDICADORES GERAIS -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #005BBF;">INDICADORES GERAIS</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:center;width:200px;">Indicador</td>
        <td style="${hd}text-align:center;width:120px;">Valor</td>
        <td style="${hd}text-align:center;width:200px;">Indicador</td>
        <td style="${hd}text-align:center;width:120px;">Valor</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr>
        <td style="${cd}background:#E3F2FD;font-weight:bold;">Total de Produtos (SKUs)</td>
        <td style="${cd}background:#E3F2FD;text-align:center;font-size:14px;font-weight:bold;color:#005BBF;">${totalProducts}</td>
        <td style="${cd}background:#E8F5E9;font-weight:bold;">Total de Unidades</td>
        <td style="${cd}background:#E8F5E9;text-align:center;font-size:14px;font-weight:bold;color:#1B5E20;">${totalUnits.toLocaleString('pt-BR')}</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr>
        <td style="${cd}background:#FFF8E1;font-weight:bold;">Itens com Estoque Baixo</td>
        <td style="${cd}background:#FFF8E1;text-align:center;font-size:14px;font-weight:bold;color:#E65100;">${lowStockCount}</td>
        <td style="${cd}background:#FCE4EC;font-weight:bold;">Requisições Pendentes</td>
        <td style="${cd}background:#FCE4EC;text-align:center;font-size:14px;font-weight:bold;color:#B71C1C;">${pendingReqs}</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 2: MOVIMENTAÇÃO -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #005BBF;">MOVIMENTAÇÃO DE ESTOQUE</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:center;">Período</td>
        <td style="${hd}text-align:center;">Entradas (Un.)</td>
        <td style="${hd}text-align:center;">Saídas (Un.)</td>
        <td style="${hd}text-align:center;">Saldo</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr>
        <td style="${cd}font-weight:bold;">Últimos 7 dias</td>
        <td style="${cd}text-align:center;color:#1B5E20;font-weight:bold;">+${entries7d.toLocaleString('pt-BR')}</td>
        <td style="${cd}text-align:center;color:#B71C1C;font-weight:bold;">-${exits7d.toLocaleString('pt-BR')}</td>
        <td style="${cd}text-align:center;font-weight:bold;color:${(entries7d - exits7d) >= 0 ? '#1B5E20' : '#B71C1C'};">${(entries7d - exits7d) >= 0 ? '+' : ''}${(entries7d - exits7d).toLocaleString('pt-BR')}</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr>
        <td style="${cd}background:#F8F9FA;font-weight:bold;">Últimos 30 dias</td>
        <td style="${cd}background:#F8F9FA;text-align:center;color:#1B5E20;font-weight:bold;">+${entries30d.toLocaleString('pt-BR')}</td>
        <td style="${cd}background:#F8F9FA;text-align:center;color:#B71C1C;font-weight:bold;">-${exits30d.toLocaleString('pt-BR')}</td>
        <td style="${cd}background:#F8F9FA;text-align:center;font-weight:bold;color:${(entries30d - exits30d) >= 0 ? '#1B5E20' : '#B71C1C'};">${(entries30d - exits30d) >= 0 ? '+' : ''}${(entries30d - exits30d).toLocaleString('pt-BR')}</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 3: DISTRIBUIÇÃO POR GRUPO -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #005BBF;">DISTRIBUIÇÃO POR GRUPO</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:left;width:200px;">Grupo</td>
        <td style="${hd}text-align:center;">Produtos</td>
        <td style="${hd}text-align:center;">Unidades</td>
        <td style="${hd}text-align:center;">Est. Baixo</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    ${groupRows}
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 4: TOP SAÍDAS 3 DIAS -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #B71C1C;">TOP SAÍDAS — ÚLTIMOS 3 DIAS</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:center;width:60px;background:#B71C1C;border-color:#8E1515;">#</td>
        <td style="${hd}text-align:left;background:#B71C1C;border-color:#8E1515;">Material</td>
        <td style="${hd}text-align:center;background:#B71C1C;border-color:#8E1515;">Código</td>
        <td style="${hd}text-align:center;background:#B71C1C;border-color:#8E1515;">Qtd Saída</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    ${renderTopTable(top3Days, '3 dias')}
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 5: TOP SAÍDAS 7 DIAS -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #E65100;">TOP SAÍDAS — ÚLTIMOS 7 DIAS</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:center;width:60px;background:#E65100;border-color:#BF4400;">#</td>
        <td style="${hd}text-align:left;background:#E65100;border-color:#BF4400;">Material</td>
        <td style="${hd}text-align:center;background:#E65100;border-color:#BF4400;">Código</td>
        <td style="${hd}text-align:center;background:#E65100;border-color:#BF4400;">Qtd Saída</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    ${renderTopTable(top7Days, '7 dias')}
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 6: TOP SAÍDAS 30 DIAS -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #795900;">TOP SAÍDAS — ÚLTIMOS 30 DIAS</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:center;width:60px;background:#795900;border-color:#5C4400;">#</td>
        <td style="${hd}text-align:left;background:#795900;border-color:#5C4400;">Material</td>
        <td style="${hd}text-align:center;background:#795900;border-color:#5C4400;">Código</td>
        <td style="${hd}text-align:center;background:#795900;border-color:#5C4400;">Qtd Saída</td>
        <td colspan="2" style="border:none;"></td>
    </tr>
    ${renderTopTable(top30Days, '30 dias')}
    <tr><td colspan="6" style="height:6px;"></td></tr>

    <!-- SEÇÃO 7: ITENS COM ESTOQUE BAIXO -->
    <tr><td colspan="6" style="${secTitle}border-bottom:2px solid #E65100;">ITENS COM ESTOQUE BAIXO (Top ${Math.min(lowStockCount, 15)})</td></tr>
    <tr><td colspan="6" style="height:4px;"></td></tr>
    <tr>
        <td style="${hd}text-align:left;width:240px;">Material</td>
        <td style="${hd}text-align:center;">Código RM</td>
        <td style="${hd}text-align:center;">Grupo</td>
        <td style="${hd}text-align:center;">Estoque</td>
        <td style="${hd}text-align:center;">Mínimo</td>
        <td style="${hd}text-align:center;background:#B71C1C;border-color:#8E1515;">Déficit</td>
    </tr>
    ${lowStockRows || `<tr><td colspan="6" style="${cd}color:#999;text-align:center;">Nenhum item com estoque baixo</td></tr>`}
    <tr><td colspan="6" style="height:10px;"></td></tr>

    <!-- RODAPÉ -->
    <tr><td colspan="6" style="border-top:2px solid #005BBF;padding:8px;font-size:10px;color:#999;font-family:Calibri,sans-serif;">Estoque UHE Estrela — Relatório gerado automaticamente em ${esc(dataGerada)}</td></tr>
</table>
</body></html>`;

            const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.setAttribute("href", url);
            link.setAttribute("download", `kpis_estoque_${timestamp}.xls`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            showToast("Planilha de KPIs gerada com sucesso!");
        });

        // Listeners da view de Sugestão de Compras
        document.addEventListener('click', (e) => {
            // Filtro de status
            if (e.target.classList.contains('compras-filter-btn')) {
                document.querySelectorAll('.compras-filter-btn').forEach(btn => {
                    btn.style.background = '';
                    btn.style.color = '';
                    btn.style.removeProperty('background');
                });
                e.target.style.background = '#191c1d';
                e.target.style.color = '#fff';
                const sel = document.getElementById('compras-period-select');
                renderComprasView(sel ? parseInt(sel.value) : 15);
            }

            // Exportar itens sugeridos para Excel (formato XLS profissional)
            if (e.target.id === 'compras-excel-btn' || e.target.closest('#compras-excel-btn')) {
                const periodLabel = (() => {
                    const sel = document.getElementById('compras-period-select');
                    return sel ? sel.options[sel.selectedIndex].text : `Últimos ${lastComprasPeriodDays} dias`;
                })();

                const toBuy = (lastComprasFilteredItems || []).filter(i => i.suggestedQty > 0);

                if (!toBuy.length) {
                    showToast('Nenhum item com compra sugerida para exportar.', true);
                    return;
                }

                const esc = v => String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
                const dataGerada = new Date().toLocaleString('pt-BR');
                const timestamp = new Date().toISOString().slice(0, 10);

                const urgencyColor = (s) => {
                    if (s === 'critico') return { bg: '#FDEDED', fg: '#B71C1C', label: 'CRÍTICO' };
                    if (s === 'atencao') return { bg: '#FFF8E1', fg: '#E65100', label: 'ATENÇÃO' };
                    return { bg: '#E3F2FD', fg: '#0D47A1', label: 'MONITORAR' };
                };

                let totalSugerido = 0;
                let criticos = 0;
                let atencao = 0;
                toBuy.forEach(i => {
                    totalSugerido += i.suggestedQty;
                    if (i.status === 'critico') criticos++;
                    else if (i.status === 'atencao') atencao++;
                });

                let rows = '';
                toBuy.forEach(({ p, dailyRate, daysRemaining, suggestedQty, status }, idx) => {
                    const u = urgencyColor(status);
                    const rowBg = idx % 2 === 0 ? '#FFFFFF' : '#F8F9FA';
                    rows += `<tr>
                        <td style="background:${u.bg};color:${u.fg};font-weight:bold;text-align:center;border:1px solid #D0D0D0;padding:6px 8px;font-size:11px;">${esc(u.label)}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;font-size:11px;">${esc(p.name)}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;">${esc(p.codeRM || p.code || '-')}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;font-size:11px;">${esc(p.group || '-')}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;">${esc(p.unit || 'UN')}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;font-size:11px;">${esc(p.location || '-')}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;">${p.quantity ?? 0}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;">${p.minQuantity ?? 0}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;">${dailyRate > 0 ? dailyRate.toFixed(2) : '0'}</td>
                        <td style="background:${rowBg};border:1px solid #D0D0D0;padding:6px 8px;text-align:center;font-size:11px;font-weight:bold;color:${daysRemaining !== null && daysRemaining <= 7 ? '#B71C1C' : '#333'};">${daysRemaining ?? '-'}</td>
                        <td style="background:#E8F5E9;border:1px solid #D0D0D0;padding:6px 10px;text-align:center;font-size:12px;font-weight:bold;color:#1B5E20;">${suggestedQty}</td>
                    </tr>`;
                });

                const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8">
<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>Compras</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
</head><body>
<table>
    <tr><td colspan="11" style="font-size:18px;font-weight:bold;color:#005BBF;padding:12px 8px 2px 8px;font-family:Calibri,sans-serif;">SUGESTÃO DE COMPRAS — ESTOQUE UHE ESTRELA</td></tr>
    <tr>
        <td colspan="3" style="font-size:11px;color:#555;padding:2px 8px;font-family:Calibri,sans-serif;">Período: ${esc(periodLabel)}</td>
        <td colspan="3" style="font-size:11px;color:#555;padding:2px 8px;font-family:Calibri,sans-serif;">Gerado em: ${esc(dataGerada)}</td>
        <td colspan="5" style="font-size:11px;color:#555;padding:2px 8px;font-family:Calibri,sans-serif;">Total de itens: ${toBuy.length}</td>
    </tr>
    <tr>
        <td colspan="3" style="font-size:11px;color:#B71C1C;padding:2px 8px;font-family:Calibri,sans-serif;">Críticos: ${criticos}</td>
        <td colspan="3" style="font-size:11px;color:#E65100;padding:2px 8px;font-family:Calibri,sans-serif;">Atenção: ${atencao}</td>
        <td colspan="5" style="font-size:11px;color:#1B5E20;padding:2px 8px;font-family:Calibri,sans-serif;">Qtd total sugerida: ${totalSugerido}</td>
    </tr>
    <tr><td colspan="11" style="height:10px;"></td></tr>
    <tr>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:90px;">Urgência</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:left;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:280px;">Material</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:100px;">Código RM</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:left;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:140px;">Grupo</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:60px;">Unid.</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:left;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:120px;">Local</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:70px;">Estoque</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:70px;">Mínimo</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:80px;">Cons./Dia</th>
        <th style="background:#005BBF;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #004A9F;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:80px;">Dias Rest.</th>
        <th style="background:#1B5E20;color:#FFFFFF;font-weight:bold;text-align:center;border:1px solid #145218;padding:8px 10px;font-size:11px;font-family:Calibri,sans-serif;width:90px;">Qtd Compra</th>
    </tr>
    ${rows}
    <tr><td colspan="11" style="height:6px;"></td></tr>
    <tr>
        <td colspan="9" style="text-align:right;font-size:11px;font-weight:bold;padding:6px 8px;border-top:2px solid #005BBF;font-family:Calibri,sans-serif;color:#333;">TOTAL A COMPRAR →</td>
        <td colspan="2" style="text-align:center;font-size:14px;font-weight:bold;padding:6px 8px;border-top:2px solid #005BBF;background:#E8F5E9;color:#1B5E20;font-family:Calibri,sans-serif;">${totalSugerido}</td>
    </tr>
</table>
</body></html>`;

                const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `compras_sugeridas_${lastComprasPeriodDays}dias_${timestamp}.xls`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                showToast('Planilha de compra exportada com sucesso.');
            }

            // Imprimir
            if (e.target.id === 'compras-print-btn' || e.target.closest('#compras-print-btn')) {
                const printArea = document.getElementById('print-area');
                const table = document.getElementById('compras-list');
                const summary = document.getElementById('compras-summary');
                if (!table) return;

                const sel = document.getElementById('compras-period-select');
                const periodo = sel ? sel.options[sel.selectedIndex].text : `Últimos ${lastComprasPeriodDays} dias`;

                printArea.innerHTML = `
                    <style>
                        @media print { body > *:not(#print-area) { display:none !important; } #print-area { display:block !important; padding:20px; font-family:sans-serif; } }
                        table { width:100%; border-collapse:collapse; font-size:12px; }
                        th, td { padding:8px 10px; border:1px solid #ddd; text-align:left; }
                        th { background:#f3f4f5; font-weight:700; }
                        h1 { font-size:18px; margin-bottom:4px; }
                        .sub { font-size:12px; color:#666; margin-bottom:16px; }
                        .summary { display:flex; gap:16px; margin-bottom:16px; }
                        .card { border:1px solid #ddd; border-radius:8px; padding:10px 16px; min-width:100px; }
                    </style>
                    <h1>Sugestão de Compras — Estoque ESTRELA</h1>
                    <p class="sub">Período: ${periodo} · Gerado em: ${new Date().toLocaleString('pt-BR')}</p>
                    ${summary.outerHTML}
                    <table>
                        <thead>
                            <tr>
                                <th>Urgência</th>
                                <th>Material</th>
                                <th>Cód. RM</th>
                                <th>Local</th>
                                <th>Estoque Atual</th>
                                <th>Estoque Mín.</th>
                                <th>Consumo/Dia</th>
                                <th>Dias Restantes</th>
                                <th>Qtd. Sugerida</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${Array.from(table.querySelectorAll('tr')).map(tr => {
                                const cells = tr.querySelectorAll('td');
                                return `<tr>${Array.from(cells).map(td => `<td>${td.innerText}</td>`).join('')}</tr>`;
                            }).join('')}
                        </tbody>
                    </table>
                `;
                printArea.classList.remove('hidden');
                window.print();
                printArea.classList.add('hidden');
            }
        });

        document.addEventListener('change', (e) => {
            if (e.target.id === 'compras-period-select') {
                renderComprasView(parseInt(e.target.value));
            }
        });

        aiDescribeBtn.addEventListener('click', async () => {
            const selectedIds = Array.from(selectedProductIds);
            if (selectedIds.length !== 1) {
                showToast("Selecione exatamente um item para ver a descrição.", true);
                return;
            }

            const productId = selectedIds[0];
            const product = products.find(p => p.id === productId);
            if (!product) {
                showToast("Produto não encontrado.", true);
                return;
            }
            
            const searchTerm = product.name;
            const imageUrl = `https://source.unsplash.com/600x400/?${encodeURIComponent(searchTerm)}`;

            const modalContent = `
                <h2 class="text-2xl font-bold mb-4">Assistente de IA</h2>
                <p class="text-slate-600 mb-4">Buscando informações para: <strong class="text-indigo-600">${searchTerm}</strong></p>
                <div id="ai-image-container" class="w-full h-48 bg-slate-200 rounded-lg flex items-center justify-center mb-4">
                    <img src="${imageUrl}" class="w-full h-48 object-cover rounded-lg" onerror="this.onerror=null;this.src='https://placehold.co/600x400/e2e8f0/475569?text=Imagem+não+encontrada';">
                </div>
                <div id="ai-description-container" class="flex justify-center p-8"><div class="spinner"></div></div>
                <div class="flex justify-end gap-4 mt-6">
                    <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 text-slate-800 rounded-lg font-semibold hover:bg-slate-300">Fechar</button>
                </div>
            `;
            document.getElementById('generic-modal').innerHTML = modalContent;
            openModal('generic-modal');

            const textPrompt = `Descreva de forma concisa e em português, em um único parágrafo, para que serve o seguinte material ou ferramenta de almoxarifado/construção: "${searchTerm}". A descrição deve ser simples e direta, ideal para um almoxarife ou trabalhador entender a sua principal aplicação.`;
            
            const description = await callGeminiAPI(textPrompt, 'gemini-1.5-flash');
            const descriptionContainer = document.getElementById('ai-description-container');
            
            if (description) {
                let audioPromise = callTTSAPI(description); 

                descriptionContainer.innerHTML = `
                    <p class="text-slate-700 bg-slate-50 p-4 rounded-lg border">${description}</p>
                    <div class="mt-4 flex justify-start">
                        <button id="play-audio-btn" class="px-4 py-2 bg-indigo-100 text-indigo-700 rounded-lg font-semibold flex items-center gap-2 hover:bg-indigo-200 transition">
                            <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15.536 8.464a5 5 0 010 7.072m2.828-9.9a9 9 0 010 12.728M5.586 15H4a1 1 0 01-1-1v-4a1 1 0 011-1h1.586l4.707-4.707C10.923 3.663 12 4.109 12 5v14c0 .89-1.077 1.337-1.707.707L5.586 15z"></path></svg>
                            Ouvir Descrição
                        </button>
                    </div>
                `;

                document.getElementById('play-audio-btn').onclick = async (e) => {
                    const btn = e.currentTarget;
                    btn.disabled = true;
                    btn.innerHTML = `<div class="spinner-small" style="border-left-color:#4f46e5;"></div>`;
                    
                    const audioData = await audioPromise;
                    if (audioData) {
                        const pcmData = base64ToArrayBuffer(audioData);
                        const pcm16 = new Int16Array(pcmData);
                        const wavBlob = pcmToWav(pcm16, 24000);
                        const audioUrl = URL.createObjectURL(wavBlob);
                        if (currentAudio) {
                            currentAudio.pause();
                        }
                        currentAudio = new Audio(audioUrl);
                        currentAudio.play();
                        currentAudio.onended = () => {
                            btn.disabled = false;
                            btn.innerHTML = `<svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15.536 8.464a5 5 0 010 7.072m2.828-9.9a9 9 0 010 12.728M5.586 15H4a1 1 0 01-1-1v-4a1 1 0 011-1h1.586l4.707-4.707C10.923 3.663 12 4.109 12 5v14c0 .89-1.077 1.337-1.707.707L5.586 15z"></path></svg>Ouvir Descrição`;
                        };
                    } else {
                        showToast("Não foi possível gerar o áudio.", true);
                        btn.disabled = false;
                        btn.innerHTML = `<svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15.536 8.464a5 5 0 010 7.072m2.828-9.9a9 9 0 010 12.728M5.586 15H4a1 1 0 01-1-1v-4a1 1 0 011-1h1.586l4.707-4.707C10.923 3.663 12 4.109 12 5v14c0 .89-1.077 1.337-1.707.707L5.586 15z"></path></svg>Ouvir Descrição`;
                    }
                };
            } else {
                descriptionContainer.innerHTML = `<p class="text-red-600 bg-red-50 p-4 rounded-lg border border-red-200">Não foi possível obter a descrição. Tente novamente.</p>`;
            }
        });

        aiExitPrompt.addEventListener('change', async (e) => {
            const promptText = e.target.value.trim();
            if (!promptText) return;

            showToast("Analisando com assistente de IA...");
            const productNames = products.map(p => `"${p.name}" (código RM: ${p.codeRM})`).join(', ');
            const locationNames = locations.map(l => `"${l.name}" (apelidos: ${(l.aliases || []).join(', ')})`).join(', ');

            const jsonSchema = {
                type: "OBJECT",
                properties: {
                    "productName": { "type": "STRING" },
                    "quantity": { "type": "NUMBER" },
                    "withdrawnBy": { "type": "STRING" },
                    "applicationLocation": { "type": "STRING" },
                    "obra": { "type": "STRING", "enum": ["TABOCA", "ESTRELA"] }
                },
                required: ["productName", "quantity", "withdrawnBy", "applicationLocation"]
            };
            
            const systemPrompt = `Você é um assistente de almoxarifado. Sua tarefa é extrair informações de uma frase e preencher um formulário em formato JSON.
            A frase descreve a retirada de um material. Você deve identificar o nome do produto, a quantidade, quem está retirando e onde o material será aplicado.
            
            Lista de produtos disponíveis: ${productNames}.
            Lista de locais de aplicação disponíveis: ${locationNames}.
            
            Selecione o produto e o local mais provável da lista. Se a obra for mencionada como 'estrela', use "ESTRELA", caso contrário, o padrão é "TABOCA".
            Responda APENAS com o JSON.`;
             const fullPrompt = `${systemPrompt}\n\nFrase do usuário: "${promptText}"`

            const config = {
                generationConfig: {
                    responseMimeType: "application/json",
                }
            };

            const resultJsonString = await callGeminiAPI(fullPrompt, 'gemini-1.5-flash', config);
            
            if (resultJsonString) {
                try {
                    const result = JSON.parse(resultJsonString.replace(/```json\n?|\n?```/g, ''));
                    
                    const foundProduct = products.find(p => p.name.toLowerCase() === result.productName?.toLowerCase() || p.codeRM === result.productName);
                    if (foundProduct) {
                        document.getElementById('exit-product-select').value = foundProduct.id;
                    } else {
                        showToast(`Produto "${result.productName}" não encontrado.`, true);
                    }

                    const foundLocation = locations.find(l => l.name.toLowerCase() === result.applicationLocation?.toLowerCase() || (l.aliases && l.aliases.some(a => a.toLowerCase() === result.applicationLocation?.toLowerCase())));
                    if (foundLocation) {
                        document.getElementById('exit-application-location').value = foundLocation.name;
                    } else {
                        document.getElementById('exit-application-location').value = result.applicationLocation || '';
                        showToast(`Local "${result.applicationLocation}" não cadastrado, mas foi preenchido.`, false);
                    }

                    document.getElementById('exit-quantity').value = result.quantity || '';
                    document.getElementById('exit-withdrawn-by').value = result.withdrawnBy || '';
                    if (result.obra) {
                        document.getElementById('exit-obra').value = result.obra;
                    }
                    
                    showToast("Formulário preenchido pela IA. Verifique os dados.", false);
                } catch (error) {
                    console.error("Error parsing AI response:", error, "Response was:", resultJsonString);
                    showToast("Não foi possível entender o pedido. Tente ser mais específico.", true);
                }
            }
        });

        searchInput.addEventListener('input', () => {
            currentPage = 1; // 🚀 Resetar para primeira página ao buscar
            renderProducts();
        });
        exitsSearchInput.addEventListener('input', (e) => renderExitLog(e.target.value));
        rmSearchInput.addEventListener('input', (e) => renderRMView(e.target.value));
        
        backupBtn.addEventListener('click', async () => {
             try {
                showLoader(true);
                const backupData = {
                    products,
                    history,
                    locations,
                    requisitions,
                    settings: appSettings
                };

                const serializedData = JSON.stringify(backupData, (key, value) => {
                    if (value && typeof value.seconds === 'number' && typeof value.nanoseconds === 'number') {
                        return { _fs_timestamp: true, seconds: value.seconds, nanoseconds: value.nanoseconds };
                    }
                    return value;
                }, 2);
                
                const blob = new Blob([serializedData], { type: 'application/json' });
                const link = document.createElement("a");
                const url = URL.createObjectURL(blob);
                const date = new Date().toISOString().slice(0, 10);
                link.setAttribute("href", url);
                link.setAttribute("download", `backup-estoque-${date}.json`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            } catch (error) {
                console.error("Erro ao criar backup:", error);
                showToast("Falha ao criar backup.", true);
            } finally {
                showLoader(false);
            }
        });

        restoreBtn.addEventListener('click', () => restoreFileInput.click());
        restoreFileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const backupData = JSON.parse(e.target.result);

                    showConfirmationModal(
                        'Restaurar Backup',
                        'Tem certeza? Esta ação substituirá TODOS os dados atuais pelos dados do arquivo de backup. Esta ação não pode ser desfeita.',
                        async () => {
                            showLoader(true);
                            try {
                                const collectionsToWipe = [
                                    productsCollectionRef, historyCollectionRef, 
                                    locationsCollectionRef, requisitionsCollectionRef
                                ];

                                for (const collRef of collectionsToWipe) {
                                    const snapshot = await getDocs(collRef);
                                    if(snapshot.docs.length > 0) {
                                        const batch = writeBatch(db);
                                        snapshot.docs.forEach(doc => batch.delete(doc.ref));
                                        await batch.commit();
                                    }
                                }
                                
                                const restoreBatch = writeBatch(db);

                                const restoreTimestamps = (data) => {
                                    return JSON.parse(JSON.stringify(data), (key, value) => {
                                        if (value && value._fs_timestamp) {
                                            return new Timestamp(value.seconds, value.nanoseconds);
                                        }
                                        return value;
                                    });
                                };
                                
                                const collectionsToRestore = {
                                    products: productsCollectionRef,
                                    history: historyCollectionRef,
                                    locations: locationsCollectionRef,
                                    requisitions: requisitionsCollectionRef
                                };
                                
                                for(const key in collectionsToRestore) {
                                    if(backupData[key] && Array.isArray(backupData[key])) {
                                        restoreTimestamps(backupData[key]).forEach(item => {
                                            const docRef = item.id ? doc(collectionsToRestore[key], item.id) : doc(collectionsToRestore[key]);
                                            const { id, ...data } = item;
                                            restoreBatch.set(docRef, data);
                                        });
                                    }
                                }
                                
                                if(backupData.settings) {
                                    restoreBatch.set(settingsDocRef, backupData.settings);
                                }
                                
                                await restoreBatch.commit();
                                showToast('Backup restaurado com sucesso! A página será atualizada.');
                                
                                setTimeout(() => window.location.reload(), 1500);

                            } catch (error) {
                                console.error('Erro ao restaurar backup:', error);
                                showToast('Falha ao restaurar backup. Verifique o console.', true);
                                showLoader(false);
                            } finally {
                                restoreFileInput.value = '';
                            }
                        }
                    );

                } catch (error) {
                    showToast('Arquivo de backup inválido.', true);
                    console.error('Erro ao ler arquivo de backup:', error);
                    restoreFileInput.value = '';
                }
            };
            reader.readAsText(file);
        });

        // 🎨 Mostrar skeleton ao carregar página
        showProductsSkeleton(true, 8);

        // Mostrar tela de login (sessão verificada acima via initializeAppSession se já logado)
        if (!localStorage.getItem('appUser')) {
            loginScreen.classList.remove('hidden');
            appContainer.classList.add('hidden');
        }
        showLoader(false);

        // =====================================================================
        // ⭐ MÓDULO COMPLETO — Controle de Estoque UHE Estrela
        // =====================================================================

        const renderEstrelaView = () => {
            renderEstrelaEstoque();
            renderEstrelaEntradas();
            renderEstrelaSaidas();
            populateEstrelaProductSelects();
        };

        // --- Abas internas ---
        document.querySelectorAll('.estrela-tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const tab = btn.dataset.estrelaTab;
                document.querySelectorAll('.estrela-tab-btn').forEach(b => {
                    b.classList.remove('active');
                    b.style.background = 'transparent';
                    b.style.color = '#414754';
                });
                btn.classList.add('active');
                btn.style.background = '#005bbf';
                btn.style.color = '#fff';
                document.querySelectorAll('.estrela-tab-content').forEach(c => c.classList.add('hidden'));
                document.getElementById(`estrela-tab-${tab}`).classList.remove('hidden');
            });
        });

        // --- Busca no estoque ---
        const estrelaSearchInput = document.getElementById('estrela-search');
        if (estrelaSearchInput) {
            estrelaSearchInput.addEventListener('input', () => renderEstrelaEstoque(estrelaSearchInput.value));
        }
        const estrelaEntrySearchInput = document.getElementById('estrela-entry-search');
        if (estrelaEntrySearchInput) {
            estrelaEntrySearchInput.addEventListener('input', () => renderEstrelaEntradas(estrelaEntrySearchInput.value));
        }
        const estrelaExitSearchInput = document.getElementById('estrela-exit-search');
        if (estrelaExitSearchInput) {
            estrelaExitSearchInput.addEventListener('input', () => renderEstrelaSaidas(estrelaExitSearchInput.value));
        }

        // --- Populate selects ---
        const populateEstrelaProductSelects = () => {
            const sorted = [...estrelaProducts].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            [document.getElementById('estrela-entry-product'), document.getElementById('estrela-exit-product')].forEach(sel => {
                if (!sel) return;
                const val = sel.value;
                sel.innerHTML = '<option value="">Selecione um produto...</option>' +
                    sorted.map(p => `<option value="${p.id}">${p.code || ''} — ${p.name} (Estoque: ${p.quantity ?? 0})</option>`).join('');
                if (val) sel.value = val;
            });
        };

        // --- Gerar código sequencial ---
        const generateEstrelaCode = () => {
            const codes = estrelaProducts.map(p => {
                const m = (p.code || '').match(/^EST-(\d+)$/);
                return m ? parseInt(m[1]) : 0;
            });
            const next = Math.max(0, ...codes) + 1;
            return `EST-${String(next).padStart(4, '0')}`;
        };

        // ============================
        // RENDER: ABA ESTOQUE
        // ============================
        const renderEstrelaEstoque = (search = '') => {
            const summaryEl = document.getElementById('estrela-summary');
            const tbody = document.getElementById('estrela-products-list');
            const noMsg = document.getElementById('no-estrela-products');
            if (!tbody) return;

            let filtered = [...estrelaProducts];
            if (search.trim()) {
                const q = search.toLowerCase();
                filtered = filtered.filter(p =>
                    (p.name || '').toLowerCase().includes(q) ||
                    (p.code || '').toLowerCase().includes(q) ||
                    (p.codeRM || '').toLowerCase().includes(q) ||
                    (p.location || '').toLowerCase().includes(q) ||
                    (p.group || '').toLowerCase().includes(q)
                );
            }
            filtered.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

            // Summary cards
            const totalItems = estrelaProducts.length;
            const totalUnits = estrelaProducts.reduce((s, p) => s + (p.quantity || 0), 0);
            const lowStock = estrelaProducts.filter(p => (p.quantity || 0) <= (p.minQuantity || 0)).length;
            const zeroStock = estrelaProducts.filter(p => (p.quantity || 0) === 0).length;

            if (summaryEl) {
                summaryEl.innerHTML = `
                    <div class="estrela-summary-card"><div class="value">${totalItems}</div><div class="label">Produtos</div></div>
                    <div class="estrela-summary-card"><div class="value">${totalUnits.toLocaleString('pt-BR')}</div><div class="label">Unidades</div></div>
                    <div class="estrela-summary-card" style="border-bottom:3px solid #fbbc05;"><div class="value" style="color:#795900;">${lowStock}</div><div class="label">Estoque Baixo</div></div>
                    <div class="estrela-summary-card" style="border-bottom:3px solid #ba1a1a;"><div class="value" style="color:#ba1a1a;">${zeroStock}</div><div class="label">Zerados</div></div>
                `;
            }

            if (filtered.length === 0) {
                tbody.innerHTML = '';
                if (noMsg) noMsg.classList.remove('hidden');
                return;
            }
            if (noMsg) noMsg.classList.add('hidden');

            tbody.innerHTML = filtered.map(p => {
                const qty = p.quantity || 0;
                const min = p.minQuantity || 0;
                let statusClass = 'estrela-status-ok', statusText = 'OK';
                if (qty === 0) { statusClass = 'estrela-status-low'; statusText = 'ZERADO'; }
                else if (qty <= min) { statusClass = 'estrela-status-warn'; statusText = 'BAIXO'; }

                return `<tr class="hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-3 font-mono text-xs font-bold" style="color:#005bbf;">${p.code || '—'}</td>
                    <td class="px-4 py-3">
                        <p class="font-semibold text-sm" style="color:#191c1d;">${p.name || '—'}</p>
                        <p class="text-xs" style="color:#727785;">RM: ${p.codeRM || '—'}</p>
                    </td>
                    <td class="px-4 py-3 text-xs">${p.group || '—'}</td>
                    <td class="px-4 py-3 text-center text-xs">${p.unit || 'Un'}</td>
                    <td class="px-4 py-3 text-center font-bold">${qty}</td>
                    <td class="px-4 py-3 text-center text-xs" style="color:#727785;">${min}</td>
                    <td class="px-4 py-3 text-xs">${p.location || '—'}</td>
                    <td class="px-4 py-3 text-center"><span class="${statusClass}">${statusText}</span></td>
                    <td class="px-4 py-3 text-center">
                        <div class="flex items-center justify-center gap-1">
                            <button onclick="editEstrelaProduct('${p.id}')" class="p-1.5 rounded-lg transition" style="background:rgba(0,91,191,0.08); color:#005bbf;" title="Editar">
                                <span class="material-symbols-outlined" style="font-size:18px;">edit</span>
                            </button>
                            <button onclick="deleteEstrelaProduct('${p.id}')" class="p-1.5 rounded-lg transition" style="background:rgba(186,26,26,0.08); color:#ba1a1a;" title="Excluir">
                                <span class="material-symbols-outlined" style="font-size:18px;">delete</span>
                            </button>
                        </div>
                    </td>
                </tr>`;
            }).join('');

            populateEstrelaProductSelects();
        };

        // ============================
        // RENDER: ABA ENTRADAS
        // ============================
        const renderEstrelaEntradas = (search = '') => {
            const tbody = document.getElementById('estrela-entries-list');
            const noMsg = document.getElementById('no-estrela-entries');
            if (!tbody) return;

            let filtered = [...estrelaEntries].sort((a, b) => {
                const da = a.date ? (a.date.seconds || 0) : 0;
                const db2 = b.date ? (b.date.seconds || 0) : 0;
                return db2 - da;
            });

            if (search.trim()) {
                const q = search.toLowerCase();
                filtered = filtered.filter(e =>
                    (e.productName || '').toLowerCase().includes(q) ||
                    (e.nf || '').toLowerCase().includes(q) ||
                    (e.supplier || '').toLowerCase().includes(q) ||
                    (e.receivedBy || '').toLowerCase().includes(q)
                );
            }

            if (filtered.length === 0) {
                tbody.innerHTML = '';
                if (noMsg) noMsg.classList.remove('hidden');
                return;
            }
            if (noMsg) noMsg.classList.add('hidden');

            tbody.innerHTML = filtered.map(e => {
                const dateStr = e.entryDate || (e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '—');
                const price = e.unitPrice ? `R$ ${parseFloat(e.unitPrice).toFixed(2)}` : '—';
                return `<tr class="hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-3 text-xs font-semibold" style="color:#006e2c;">${dateStr}</td>
                    <td class="px-4 py-3 text-sm font-semibold">${e.productName || '—'}</td>
                    <td class="px-4 py-3 text-center font-bold" style="color:#006e2c;">+${e.quantity || 0}</td>
                    <td class="px-4 py-3 text-xs">${e.supplier || '—'}</td>
                    <td class="px-4 py-3 text-xs font-mono">${e.nf || '—'}</td>
                    <td class="px-4 py-3 text-xs">${price}</td>
                    <td class="px-4 py-3 text-xs">${e.receivedBy || '—'}</td>
                    <td class="px-4 py-3 text-xs" style="max-width:180px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${(e.observation || '').replace(/"/g, '&quot;')}">${e.observation || '—'}</td>
                </tr>`;
            }).join('');
        };

        // ============================
        // RENDER: ABA SAÍDAS
        // ============================
        const renderEstrelaSaidas = (search = '') => {
            const tbody = document.getElementById('estrela-exits-list');
            const noMsg = document.getElementById('no-estrela-exits');
            if (!tbody) return;

            let filtered = [...estrelaExits].sort((a, b) => {
                const da = a.date ? (a.date.seconds || 0) : 0;
                const db2 = b.date ? (b.date.seconds || 0) : 0;
                return db2 - da;
            });

            if (search.trim()) {
                const q = search.toLowerCase();
                filtered = filtered.filter(e =>
                    (e.productName || '').toLowerCase().includes(q) ||
                    (e.who || '').toLowerCase().includes(q) ||
                    (e.leader || '').toLowerCase().includes(q) ||
                    (e.applicationLocation || '').toLowerCase().includes(q) ||
                    (e.os || '').toLowerCase().includes(q)
                );
            }

            if (filtered.length === 0) {
                tbody.innerHTML = '';
                if (noMsg) noMsg.classList.remove('hidden');
                return;
            }
            if (noMsg) noMsg.classList.add('hidden');

            tbody.innerHTML = filtered.map(e => {
                const dateStr = e.exitDate || (e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '—');
                return `<tr class="hover:bg-slate-50 transition-colors">
                    <td class="px-4 py-3 text-xs font-semibold" style="color:#ba1a1a;">${dateStr}</td>
                    <td class="px-4 py-3 text-sm font-semibold">${e.productName || '—'}</td>
                    <td class="px-4 py-3 text-center font-bold" style="color:#ba1a1a;">-${e.quantity || 0}</td>
                    <td class="px-4 py-3 text-xs">${e.who || '—'}</td>
                    <td class="px-4 py-3 text-xs">${e.leader || '—'}</td>
                    <td class="px-4 py-3 text-xs">${e.applicationLocation || '—'}</td>
                    <td class="px-4 py-3 text-xs font-mono">${e.os || '—'}</td>
                    <td class="px-4 py-3 text-xs" style="max-width:180px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${(e.observation || '').replace(/"/g, '&quot;')}">${e.observation || '—'}</td>
                </tr>`;
            }).join('');
        };

        // ============================
        // CRUD: Adicionar produto UHE Estrela
        // ============================
        const addEstrelaProductBtn = document.getElementById('estrela-add-product-btn');
        if (addEstrelaProductBtn) {
            addEstrelaProductBtn.addEventListener('click', () => {
                const modal = document.getElementById('generic-modal');
                modal.innerHTML = `
                    <h2 class="text-xl font-bold mb-4" style="color:#191c1d;"><span class="material-symbols-outlined align-middle mr-1" style="font-size:24px; color:#fbbc05;">star</span> Novo Produto — UHE Estrela</h2>
                    <form id="estrela-add-form" class="space-y-3">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Nome do Produto</label>
                            <input type="text" id="ep-name" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Código RM</label>
                                <input type="text" id="ep-code-rm" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                            </div>
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Código Interno</label>
                                <input type="text" id="ep-code" class="w-full p-2.5 border border-slate-200 rounded-lg bg-slate-100 text-sm" value="${generateEstrelaCode()}" readonly>
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Grupo</label>
                                <select id="ep-group" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                                    <option>Elétrico</option><option>Hidráulico</option><option>Consumível</option>
                                    <option>Ferramentas Manuais</option><option>Material de Corte e Solda</option>
                                    <option>Escritório</option><option>Segurança</option><option>Outros</option>
                                </select>
                            </div>
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Unidade</label>
                                <select id="ep-unit" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                                    <option>Unidade</option><option>Peça</option><option>Metros</option>
                                    <option>Litro</option><option>Quilo</option><option>Caixa</option>
                                    <option>Pacote</option><option>Rolo</option><option>Galão</option><option>Saco</option>
                                </select>
                            </div>
                        </div>
                        <div class="grid grid-cols-2 gap-3">
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Quantidade Inicial</label>
                                <input type="number" id="ep-qty" min="0" value="0" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                            </div>
                            <div>
                                <label class="block text-sm font-medium text-slate-700 mb-1">Quantidade Mínima</label>
                                <input type="number" id="ep-min" min="0" value="0" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                            </div>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Localização</label>
                            <input type="text" id="ep-location" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" placeholder="Ex: Prateleira A-01" required>
                        </div>
                        <div class="flex justify-end gap-2 pt-3">
                            <button type="button" id="ep-cancel" class="px-4 py-2.5 rounded-lg text-sm font-semibold" style="background:#f3f4f5; color:#414754;">Cancelar</button>
                            <button type="submit" class="px-4 py-2.5 rounded-lg text-sm font-bold text-white" style="background:#005bbf;">Cadastrar</button>
                        </div>
                    </form>
                `;
                openModal('generic-modal');

                document.getElementById('ep-cancel').onclick = () => closeGenericModal();

                document.getElementById('estrela-add-form').addEventListener('submit', async (e) => {
                    e.preventDefault();
                    const name = document.getElementById('ep-name').value.trim();
                    if (!name) return showToast('Informe o nome do produto.', true);
                    try {
                        await addDoc(estrelaProductsRef, {
                            name,
                            code: document.getElementById('ep-code').value.trim(),
                            codeRM: document.getElementById('ep-code-rm').value.trim(),
                            group: document.getElementById('ep-group').value,
                            unit: document.getElementById('ep-unit').value,
                            quantity: parseInt(document.getElementById('ep-qty').value) || 0,
                            minQuantity: parseInt(document.getElementById('ep-min').value) || 0,
                            location: document.getElementById('ep-location').value.trim(),
                            createdAt: serverTimestamp(),
                            createdBy: currentUser?.displayName || 'Anônimo'
                        });
                        showToast('⭐ Produto cadastrado na UHE Estrela!');
                        closeGenericModal();
                    } catch (err) {
                        console.error(err);
                        showToast('Erro ao cadastrar produto.', true);
                    }
                });
            });
        }

        // ============================
        // CRUD: Editar produto UHE Estrela
        // ============================
        window.editEstrelaProduct = (id) => {
            const p = estrelaProducts.find(x => x.id === id);
            if (!p) return;
            const modal = document.getElementById('generic-modal');
            modal.innerHTML = `
                <h2 class="text-xl font-bold mb-4" style="color:#191c1d;">Editar Produto — UHE Estrela</h2>
                <form id="estrela-edit-form" class="space-y-3">
                    <div>
                        <label class="block text-sm font-medium text-slate-700 mb-1">Nome do Produto</label>
                        <input type="text" id="epe-name" value="${(p.name || '').replace(/"/g, '&quot;')}" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Código RM</label>
                            <input type="text" id="epe-code-rm" value="${p.codeRM || ''}" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Código Interno</label>
                            <input type="text" id="epe-code" value="${p.code || ''}" class="w-full p-2.5 border border-slate-200 rounded-lg bg-slate-100 text-sm" readonly>
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Grupo</label>
                            <select id="epe-group" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                                ${['Elétrico','Hidráulico','Consumível','Ferramentas Manuais','Material de Corte e Solda','Escritório','Segurança','Outros'].map(g => `<option ${g===p.group?'selected':''}>${g}</option>`).join('')}
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Unidade</label>
                            <select id="epe-unit" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm">
                                ${['Unidade','Peça','Metros','Litro','Quilo','Caixa','Pacote','Rolo','Galão','Saco'].map(u => `<option ${u===p.unit?'selected':''}>${u}</option>`).join('')}
                            </select>
                        </div>
                    </div>
                    <div class="grid grid-cols-2 gap-3">
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Quantidade</label>
                            <input type="number" id="epe-qty" min="0" value="${p.quantity || 0}" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                        </div>
                        <div>
                            <label class="block text-sm font-medium text-slate-700 mb-1">Quantidade Mínima</label>
                            <input type="number" id="epe-min" min="0" value="${p.minQuantity || 0}" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                        </div>
                    </div>
                    <div>
                        <label class="block text-sm font-medium text-slate-700 mb-1">Localização</label>
                        <input type="text" id="epe-location" value="${p.location || ''}" class="w-full p-2.5 border border-slate-200 rounded-lg text-sm" required>
                    </div>
                    <div class="flex justify-end gap-2 pt-3">
                        <button type="button" id="epe-cancel" class="px-4 py-2.5 rounded-lg text-sm font-semibold" style="background:#f3f4f5; color:#414754;">Cancelar</button>
                        <button type="submit" class="px-4 py-2.5 rounded-lg text-sm font-bold text-white" style="background:#005bbf;">Salvar</button>
                    </div>
                </form>
            `;
            openModal('generic-modal');

            document.getElementById('epe-cancel').onclick = () => closeGenericModal();

            document.getElementById('estrela-edit-form').addEventListener('submit', async (ev) => {
                ev.preventDefault();
                try {
                    await updateDoc(doc(estrelaProductsRef, id), {
                        name: document.getElementById('epe-name').value.trim(),
                        codeRM: document.getElementById('epe-code-rm').value.trim(),
                        group: document.getElementById('epe-group').value,
                        unit: document.getElementById('epe-unit').value,
                        quantity: parseInt(document.getElementById('epe-qty').value) || 0,
                        minQuantity: parseInt(document.getElementById('epe-min').value) || 0,
                        location: document.getElementById('epe-location').value.trim(),
                        updatedAt: serverTimestamp(),
                        updatedBy: currentUser?.displayName || 'Anônimo'
                    });
                    showToast('⭐ Produto atualizado!');
                    closeGenericModal();
                } catch (err) {
                    console.error(err);
                    showToast('Erro ao atualizar produto.', true);
                }
            });
        };

        // ============================
        // CRUD: Excluir produto UHE Estrela
        // ============================
        window.deleteEstrelaProduct = (id) => {
            const p = estrelaProducts.find(x => x.id === id);
            if (!p) return;
            const modal = document.getElementById('generic-modal');
            modal.innerHTML = `
                <h2 class="text-xl font-bold mb-3" style="color:#ba1a1a;">Confirmar Exclusão</h2>
                <p class="text-sm mb-4" style="color:#414754;">Deseja excluir <strong>${p.name}</strong> do estoque UHE Estrela? Esta ação não pode ser desfeita.</p>
                <div class="flex justify-end gap-2">
                    <button id="edel-cancel" class="px-4 py-2.5 rounded-lg text-sm font-semibold" style="background:#f3f4f5; color:#414754;">Cancelar</button>
                    <button id="edel-confirm" class="px-4 py-2.5 rounded-lg text-sm font-bold text-white" style="background:#ba1a1a;">Excluir</button>
                </div>
            `;
            openModal('generic-modal');

            document.getElementById('edel-cancel').onclick = () => closeGenericModal();
            document.getElementById('edel-confirm').onclick = async () => {
                try {
                    await deleteDoc(doc(estrelaProductsRef, id));
                    showToast('Produto excluído do estoque Estrela.');
                    closeGenericModal();
                } catch (err) {
                    console.error(err);
                    showToast('Erro ao excluir.', true);
                }
            };
        };

        // ============================
        // FORM: Registrar entrada UHE Estrela
        // ============================
        const estrelaEntryForm = document.getElementById('estrela-entry-form');
        if (estrelaEntryForm) {
            estrelaEntryForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const productId = document.getElementById('estrela-entry-product').value;
                const qty = parseInt(document.getElementById('estrela-entry-qty').value) || 0;
                const entryDate = document.getElementById('estrela-entry-date').value;
                const supplier = document.getElementById('estrela-entry-supplier').value.trim();
                const nf = document.getElementById('estrela-entry-nf').value.trim();

                if (!productId) return showToast('Selecione um produto.', true);
                if (qty <= 0) return showToast('Quantidade deve ser maior que zero.', true);
                if (!entryDate) return showToast('Informe a data de chegada.', true);
                if (!supplier) return showToast('Informe o fornecedor.', true);
                if (!nf) return showToast('Informe o número da NF.', true);

                const product = estrelaProducts.find(p => p.id === productId);
                if (!product) return showToast('Produto não encontrado.', true);

                try {
                    const newQty = (product.quantity || 0) + qty;
                    await updateDoc(doc(estrelaProductsRef, productId), { quantity: newQty, updatedAt: serverTimestamp() });
                    await addDoc(estrelaEntriesRef, {
                        productId,
                        productName: product.name,
                        productCode: product.code || '',
                        quantity: qty,
                        entryDate,
                        supplier,
                        nf,
                        unitPrice: document.getElementById('estrela-entry-price').value || null,
                        receivedBy: document.getElementById('estrela-entry-received-by').value.trim() || null,
                        observation: document.getElementById('estrela-entry-obs').value.trim() || null,
                        date: serverTimestamp(),
                        registeredBy: currentUser?.displayName || 'Anônimo'
                    });
                    showToast(`⭐ Entrada registrada: +${qty} ${product.name}`);
                    estrelaEntryForm.reset();
                    document.getElementById('estrela-entry-date').valueAsDate = new Date();
                } catch (err) {
                    console.error(err);
                    showToast('Erro ao registrar entrada.', true);
                }
            });
        }

        // ============================
        // FORM: Registrar saída UHE Estrela
        // ============================
        const estrelaExitForm = document.getElementById('estrela-exit-form');
        if (estrelaExitForm) {
            estrelaExitForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const productId = document.getElementById('estrela-exit-product').value;
                const qty = parseInt(document.getElementById('estrela-exit-qty').value) || 0;
                const exitDate = document.getElementById('estrela-exit-date').value;
                const who = document.getElementById('estrela-exit-who').value.trim();
                const leader = document.getElementById('estrela-exit-leader').value.trim();
                const appLocation = document.getElementById('estrela-exit-location').value.trim();

                if (!productId) return showToast('Selecione um produto.', true);
                if (qty <= 0) return showToast('Quantidade deve ser maior que zero.', true);
                if (!exitDate) return showToast('Informe a data da saída.', true);
                if (!who) return showToast('Informe quem pegou o material.', true);
                if (!leader) return showToast('Informe o líder/encarregado.', true);
                if (!appLocation) return showToast('Informe o local de aplicação.', true);

                const product = estrelaProducts.find(p => p.id === productId);
                if (!product) return showToast('Produto não encontrado.', true);
                if (qty > (product.quantity || 0)) return showToast(`Estoque insuficiente! Disponível: ${product.quantity || 0}`, true);

                try {
                    const newQty = (product.quantity || 0) - qty;
                    await updateDoc(doc(estrelaProductsRef, productId), { quantity: newQty, updatedAt: serverTimestamp() });
                    await addDoc(estrelaExitsRef, {
                        productId,
                        productName: product.name,
                        productCode: product.code || '',
                        quantity: qty,
                        exitDate,
                        who,
                        leader,
                        applicationLocation: appLocation,
                        os: document.getElementById('estrela-exit-os').value.trim() || null,
                        observation: document.getElementById('estrela-exit-obs').value.trim() || null,
                        date: serverTimestamp(),
                        registeredBy: currentUser?.displayName || 'Anônimo'
                    });
                    showToast(`⭐ Saída registrada: -${qty} ${product.name}`);
                    estrelaExitForm.reset();
                    document.getElementById('estrela-exit-date').valueAsDate = new Date();
                } catch (err) {
                    console.error(err);
                    showToast('Erro ao registrar saída.', true);
                }
            });
        }

        // ============================
        // Exportar CSV UHE Estrela
        // ============================
        const estrelaExportBtn = document.getElementById('estrela-export-btn');
        if (estrelaExportBtn) {
            estrelaExportBtn.addEventListener('click', () => {
                const activeTab = document.querySelector('.estrela-tab-btn.active');
                const tab = activeTab ? activeTab.dataset.estrelaTab : 'estoque';
                let csvContent = '';
                let filename = '';

                if (tab === 'estoque') {
                    csvContent = 'Código,Código RM,Produto,Grupo,Unidade,Qtd Atual,Qtd Mínima,Localização,Status\n';
                    estrelaProducts.forEach(p => {
                        const qty = p.quantity || 0;
                        const min = p.minQuantity || 0;
                        let status = 'OK';
                        if (qty === 0) status = 'ZERADO';
                        else if (qty <= min) status = 'BAIXO';
                        csvContent += `"${p.code || ''}","${p.codeRM || ''}","${(p.name || '').replace(/"/g, '""')}","${p.group || ''}","${p.unit || 'Un'}",${qty},${min},"${p.location || ''}","${status}"\n`;
                    });
                    filename = `Estoque_UHE_Estrela_${new Date().toISOString().slice(0, 10)}.csv`;
                } else if (tab === 'entrada') {
                    csvContent = 'Data,Produto,Quantidade,Fornecedor,NF,Valor Unitário,Recebido Por,Observação\n';
                    estrelaEntries.forEach(e => {
                        const dateStr = e.entryDate || '';
                        csvContent += `"${dateStr}","${(e.productName || '').replace(/"/g, '""')}",${e.quantity || 0},"${(e.supplier || '').replace(/"/g, '""')}","${e.nf || ''}","${e.unitPrice || ''}","${e.receivedBy || ''}","${(e.observation || '').replace(/"/g, '""')}"\n`;
                    });
                    filename = `Entradas_UHE_Estrela_${new Date().toISOString().slice(0, 10)}.csv`;
                } else {
                    csvContent = 'Data,Produto,Quantidade,Quem Pegou,Líder,Local Aplicação,OS,Observação\n';
                    estrelaExits.forEach(e => {
                        const dateStr = e.exitDate || '';
                        csvContent += `"${dateStr}","${(e.productName || '').replace(/"/g, '""')}",${e.quantity || 0},"${(e.who || '').replace(/"/g, '""')}","${(e.leader || '').replace(/"/g, '""')}","${(e.applicationLocation || '').replace(/"/g, '""')}","${e.os || ''}","${(e.observation || '').replace(/"/g, '""')}"\n`;
                    });
                    filename = `Saidas_UHE_Estrela_${new Date().toISOString().slice(0, 10)}.csv`;
                }

                const BOM = '\uFEFF';
                const blob = new Blob([BOM + csvContent], { type: 'text/csv;charset=utf-8;' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                a.click();
                URL.revokeObjectURL(url);
                showToast(`CSV exportado: ${filename}`);
            });
        }

        // Inicializar datas nos formulários de entrada/saída
        const estrelaEntryDateInput = document.getElementById('estrela-entry-date');
        const estrelaExitDateInput = document.getElementById('estrela-exit-date');
        if (estrelaEntryDateInput) estrelaEntryDateInput.valueAsDate = new Date();
        if (estrelaExitDateInput) estrelaExitDateInput.valueAsDate = new Date();

        // Fechar modal genérico (wrapper para módulo Estrela)
        const closeGenericModal = () => closeModal('generic-modal');

        // ============================
        // ⭐ BAIXAR PLANILHA COMPLETA UHE ESTRELA (Excel com 3 abas — dados do app)
        // ============================
        window.downloadEstrelaExcel = () => {
            const XLS = typeof XLSX !== 'undefined' ? XLSX : null;
            if (!XLS) {
                showToast('Biblioteca de planilhas não carregou. Recarregue a página.', true);
                return;
            }

            const wb = XLS.utils.book_new();
            const today = new Date().toLocaleDateString('pt-BR');
            const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

            // ─── ESTILOS ───
            const titleStyle = {
                font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '005BBF' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: { bottom: { style: 'medium', color: { rgb: '003D8A' } } }
            };
            const subtitleStyle = {
                font: { italic: true, sz: 10, color: { rgb: '555555' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
            const headerStyle = {
                font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '1A73E8' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                    top: { style: 'thin', color: { rgb: '003D8A' } },
                    bottom: { style: 'thin', color: { rgb: '003D8A' } },
                    left: { style: 'thin', color: { rgb: '003D8A' } },
                    right: { style: 'thin', color: { rgb: '003D8A' } }
                }
            };
            const headerGreen = { ...headerStyle, fill: { fgColor: { rgb: '006E2C' } }, border: { ...headerStyle.border, top: { style: 'thin', color: { rgb: '004D1A' } }, bottom: { style: 'thin', color: { rgb: '004D1A' } }, left: { style: 'thin', color: { rgb: '004D1A' } }, right: { style: 'thin', color: { rgb: '004D1A' } } } };
            const headerRed = { ...headerStyle, fill: { fgColor: { rgb: 'BA1A1A' } }, border: { ...headerStyle.border, top: { style: 'thin', color: { rgb: '8A0000' } }, bottom: { style: 'thin', color: { rgb: '8A0000' } }, left: { style: 'thin', color: { rgb: '8A0000' } }, right: { style: 'thin', color: { rgb: '8A0000' } } } };

            const cellBorder = {
                top: { style: 'thin', color: { rgb: 'D0D0D0' } },
                bottom: { style: 'thin', color: { rgb: 'D0D0D0' } },
                left: { style: 'thin', color: { rgb: 'D0D0D0' } },
                right: { style: 'thin', color: { rgb: 'D0D0D0' } }
            };
            const cellBase = { font: { sz: 10 }, alignment: { vertical: 'center' }, border: cellBorder };
            const cellCenter = { ...cellBase, alignment: { horizontal: 'center', vertical: 'center' } };
            const cellBold = { ...cellCenter, font: { sz: 10, bold: true } };
            const cellEven = { ...cellBase, fill: { fgColor: { rgb: 'F2F7FF' } } };
            const cellEvenCenter = { ...cellCenter, fill: { fgColor: { rgb: 'F2F7FF' } } };
            const cellEvenBold = { ...cellBold, fill: { fgColor: { rgb: 'F2F7FF' } } };
            const cellEvenGreen = { ...cellBase, fill: { fgColor: { rgb: 'F0FFF4' } } };
            const cellEvenGreenCenter = { ...cellCenter, fill: { fgColor: { rgb: 'F0FFF4' } } };
            const cellEvenRed = { ...cellBase, fill: { fgColor: { rgb: 'FFF5F5' } } };
            const cellEvenRedCenter = { ...cellCenter, fill: { fgColor: { rgb: 'FFF5F5' } } };

            const statusOk = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: '006E2C' } }, fill: { fgColor: { rgb: 'DCFCE7' } }, border: cellBorder };
            const statusBaixo = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: '795900' } }, fill: { fgColor: { rgb: 'FEF9C3' } }, border: cellBorder };
            const statusZerado = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: 'BA1A1A' } }, fill: { fgColor: { rgb: 'FEE2E2' } }, border: cellBorder };
            const statusOkEven = { ...statusOk, fill: { fgColor: { rgb: 'BBF7D0' } } };
            const statusBaixoEven = { ...statusBaixo, fill: { fgColor: { rgb: 'FDE68A' } } };
            const statusZeradoEven = { ...statusZerado, fill: { fgColor: { rgb: 'FECACA' } } };

            const qtyGreen = { ...cellCenter, font: { sz: 10, bold: true, color: { rgb: '006E2C' } }, border: cellBorder };
            const qtyRed = { ...cellCenter, font: { sz: 10, bold: true, color: { rgb: 'BA1A1A' } }, border: cellBorder };
            const qtyGreenEven = { ...qtyGreen, fill: { fgColor: { rgb: 'F0FFF4' } } };
            const qtyRedEven = { ...qtyRed, fill: { fgColor: { rgb: 'FFF5F5' } } };

            const applyStyles = (ws, startRow, dataLen, colCount, styleFn) => {
                for (let r = 0; r < dataLen; r++) {
                    for (let c = 0; c < colCount; c++) {
                        const addr = XLS.utils.encode_cell({ r: startRow + r, c });
                        if (ws[addr]) ws[addr].s = styleFn(r, c, ws[addr].v);
                    }
                }
            };

            // ─────────────────────────────────────────────
            // ABA 1: ESTOQUE (dados reais do app — products)
            // ─────────────────────────────────────────────
            const estoqueHeaders = ['Nº', 'Código SKU', 'Código RM', 'Produto', 'Grupo', 'Unidade', 'Qtd. Atual', 'Qtd. Mínima', 'Localização', 'Status'];
            const sortedProducts = [...products].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            const estoqueRows = sortedProducts.map((p, i) => {
                const qty = p.quantity || 0;
                const min = p.minQuantity || 0;
                let status = '✅ OK';
                if (qty === 0) status = '🔴 ZERADO';
                else if (qty <= min) status = '🟡 BAIXO';
                return [i + 1, p.code || '', p.codeRM || '', p.name || '', p.group || '', p.unit || 'Un', qty, min, p.location || '', status];
            });

            const totalUnits = products.reduce((s, p) => s + (p.quantity || 0), 0);
            const lowCount = products.filter(p => (p.quantity || 0) > 0 && (p.quantity || 0) <= (p.minQuantity || 0)).length;
            const zeroCount = products.filter(p => (p.quantity || 0) === 0).length;

            const wsE = XLS.utils.aoa_to_sheet([
                ['CONTROLE DE ESTOQUE — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total de Produtos: ${products.length} | Unidades: ${totalUnits.toLocaleString('pt-BR')} | Estoque Baixo: ${lowCount} | Zerados: ${zeroCount}`],
                [],
                estoqueHeaders,
                ...estoqueRows
            ]);
            wsE['!cols'] = [
                { wch: 5 }, { wch: 14 }, { wch: 14 }, { wch: 42 }, { wch: 22 },
                { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 24 }, { wch: 14 }
            ];
            wsE['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
            ];
            wsE['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            // Estilo título e subtítulo
            wsE['A1'].s = titleStyle;
            wsE['A2'].s = subtitleStyle;
            // Estilo cabeçalho
            for (let c = 0; c < estoqueHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsE[addr]) wsE[addr].s = headerStyle;
            }
            // Estilo dados
            applyStyles(wsE, 4, estoqueRows.length, estoqueHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 9) {
                    if (String(v).includes('ZERADO')) return isEven ? statusZeradoEven : statusZerado;
                    if (String(v).includes('BAIXO')) return isEven ? statusBaixoEven : statusBaixo;
                    return isEven ? statusOkEven : statusOk;
                }
                if (c === 0 || c === 5 || c === 6 || c === 7) return isEven ? cellEvenCenter : cellCenter;
                if (c === 1 || c === 2) return isEven ? cellEvenBold : cellBold;
                return isEven ? cellEven : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsE, 'Estoque');

            // ─────────────────────────────────────────────
            // ABA 2: ENTRADAS (dados reais do app — history filtrado)
            // ─────────────────────────────────────────────
            const entradas = history
                .filter(h => ['Entrada', 'Entrada por NF', 'Ajuste Entrada', 'Criação', 'Importação'].includes(h.type))
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            const entradaHeaders = ['Nº', 'Data', 'Tipo', 'Produto', 'Código RM', 'Quantidade', 'Novo Saldo', 'NF', 'Fornecedor', 'Observação'];
            const entradaRows = entradas.map((e, i) => {
                const dateStr = e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '';
                return [
                    i + 1, dateStr, e.type || '', e.productName || '', e.productCodeRM || '',
                    e.quantity || 0, e.newTotal ?? '', e.nfNumber || e.details || '',
                    e.supplier || '', e.observation || e.details || e.performedBy || ''
                ];
            });

            const titleGreenStyle = { ...titleStyle, fill: { fgColor: { rgb: '006E2C' } }, border: { bottom: { style: 'medium', color: { rgb: '004D1A' } } } };
            const wsEnt = XLS.utils.aoa_to_sheet([
                ['REGISTRO DE ENTRADAS — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total de registros: ${entradas.length}`],
                [],
                entradaHeaders,
                ...entradaRows
            ]);
            wsEnt['!cols'] = [
                { wch: 5 }, { wch: 12 }, { wch: 18 }, { wch: 42 }, { wch: 14 },
                { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 28 }, { wch: 20 }
            ];
            wsEnt['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
            ];
            wsEnt['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            wsEnt['A1'].s = titleGreenStyle;
            wsEnt['A2'].s = subtitleStyle;
            for (let c = 0; c < entradaHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsEnt[addr]) wsEnt[addr].s = headerGreen;
            }
            applyStyles(wsEnt, 4, entradaRows.length, entradaHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 5) return isEven ? qtyGreenEven : qtyGreen;
                if (c === 0 || c === 1 || c === 6) return isEven ? cellEvenGreenCenter : cellCenter;
                return isEven ? cellEvenGreen : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsEnt, 'Entradas');

            // ─────────────────────────────────────────────
            // ABA 3: SAÍDAS (dados reais do app — history filtrado)
            // ─────────────────────────────────────────────
            const saidas = history
                .filter(h => ['Saída', 'Saída por Requisição', 'Ajuste Saída'].includes(h.type))
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            const saidaHeaders = ['Nº', 'Data', 'Produto', 'Código RM', 'Quantidade', 'Novo Saldo', 'Retirado Por', 'Líder', 'Local Aplicação', 'Obra', 'Registrado Por'];
            const saidaRows = saidas.map((e, i) => {
                const dateStr = e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '';
                return [
                    i + 1, dateStr, e.productName || '', e.productCodeRM || '',
                    e.quantity || 0, e.newTotal ?? '', e.withdrawnBy || '',
                    e.teamLeader || '', e.applicationLocation || '', e.obra || '', e.performedBy || ''
                ];
            });

            const titleRedStyle = { ...titleStyle, fill: { fgColor: { rgb: 'BA1A1A' } }, border: { bottom: { style: 'medium', color: { rgb: '8A0000' } } } };
            const wsSai = XLS.utils.aoa_to_sheet([
                ['REGISTRO DE SAÍDAS — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total de registros: ${saidas.length}`],
                [],
                saidaHeaders,
                ...saidaRows
            ]);
            wsSai['!cols'] = [
                { wch: 5 }, { wch: 12 }, { wch: 42 }, { wch: 14 }, { wch: 12 },
                { wch: 12 }, { wch: 22 }, { wch: 18 }, { wch: 28 }, { wch: 12 }, { wch: 20 }
            ];
            wsSai['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 10 } }
            ];
            wsSai['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            wsSai['A1'].s = titleRedStyle;
            wsSai['A2'].s = subtitleStyle;
            for (let c = 0; c < saidaHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsSai[addr]) wsSai[addr].s = headerRed;
            }
            applyStyles(wsSai, 4, saidaRows.length, saidaHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 4) return isEven ? qtyRedEven : qtyRed;
                if (c === 0 || c === 1 || c === 5 || c === 9) return isEven ? cellEvenRedCenter : cellCenter;
                return isEven ? cellEvenRed : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsSai, 'Saídas');

            // ─── GERAR E BAIXAR ───
            const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
            const filename = `Controle_Estoque_UHE_Estrela_${stamp}.xlsx`;
            XLS.writeFile(wb, filename);
            showToast(`⭐ Planilha baixada: ${filename}`);
        };
