// Importações do Firebase SDK
        import { initializeApp } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-app.js";
        import { getFirestore, initializeFirestore, persistentLocalCache, persistentMultipleTabManager, collection, doc, addDoc, getDocs, setDoc, updateDoc, deleteDoc, deleteField, onSnapshot, serverTimestamp, runTransaction, writeBatch, Timestamp, getDoc } from "https://www.gstatic.com/firebasejs/11.6.1/firebase-firestore.js";

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
        /** PIN só para evitar exclusão acidental no estoque (não é segurança forte). */
        const STOCK_DELETE_PIN = '0000';
        const app = initializeApp(firebaseConfig);
        /** Cache local (IndexedDB): menos leituras cobradas ao recarregar ou reabrir o app quando os dados já estão em cache. */
        let db;
        try {
            db = initializeFirestore(app, {
                localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() })
            });
        } catch (e) {
            console.warn('Firestore: cache persistente indisponível, usando modo padrão.', e?.message || e);
            db = getFirestore(app);
        }
        // Sem Firebase Auth — autenticação gerenciada pelo próprio app

        // --- Estado da Aplicação ---
        let products = [];
        let history = [];
        let requisitions = [];
        let toolLoans = [];
        let toolLoanQueue = [];
        let locations = [];
        let appSettings = { appName: 'UHE Estrela', logoUrl: null };
        let currentProductId = null;
        let currentLocationId = null;
        let productsCollectionRef;
        let historyCollectionRef;
        let requisitionsCollectionRef;
        let toolLoansCollectionRef;
        let locationsCollectionRef;
        let settingsDocRef;
        let usersCollectionRef;
        let inventoryFilter = 'all';
        let inventorySortOrder = 'default';
        let inventoryLocationFilter = 'all';
        let isDataLoaded = false;
        let currentAudio = null;
        let selectedProductIds = new Set();
        let currentViewId = 'dashboard-view';
        let coreUnsubscribers = [];
        let estrelaUnsubscribers = [];
        
        // 🔐 Sistema de Autenticação e Permissões
        let currentUser = null;
        let userRole = null; // 'admin', 'operador', 'visualizador'
        let currentObraId = null; // ID da obra selecionada para multi-tenancy
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
        /** Snapshot da última análise (export Excel independente dos filtros da tabela). */
        let lastComprasExport = {
            coverageDays: 30,
            periodDays: 15,
            pedido: [],
            turnover: [],
            zero: [],
            low: []
        };
        
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
        const activityLogList = document.getElementById('activity-log-list');
        const activityLogSearchInput = document.getElementById('activity-log-search-input');
        const noActivityLogMessage = document.getElementById('no-activity-log-message');
        const sortOrderSelect = document.getElementById('sort-order');
        const locationFilterSelect = document.getElementById('location-filter');
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
        const printSelectedReqsBtn = document.getElementById('print-selected-reqs-btn');
        const generateExitReportBtn = document.getElementById('generate-exit-report-btn');
        const generateConsumptionReportBtn = document.getElementById('generate-consumption-report-btn');
        const generateKpiReportBtn = document.getElementById('generate-kpi-report-btn');
        const addLocationForm = document.getElementById('add-location-form');
        const loginUsernameInput = document.getElementById('login-username');
        const loginPasswordInput = document.getElementById('login-password');
        const loginObraSelect = document.getElementById('login-obra');
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
        const toolLoanForm = document.getElementById('tool-loan-form');
        const toolLoanSearchInput = document.getElementById('tool-loan-search');
        const toolLoanProductSelect = document.getElementById('tool-loan-product');
        const toolLoansOpenList = document.getElementById('tool-loans-open-list');
        const toolLoansReturnedList = document.getElementById('tool-loans-returned-list');
        const noToolLoansOpenMessage = document.getElementById('no-tool-loans-open-message');
        const noToolLoansReturnedMessage = document.getElementById('no-tool-loans-returned-message');
        const exportToolLoansBtn = document.getElementById('export-tool-loans-btn');
        const toolLoanAddQueueBtn = document.getElementById('tool-loan-add-queue-btn');
        const toolLoanQueueBody = document.getElementById('tool-loan-queue-body');
        const toolLoanQueueEmpty = document.getElementById('tool-loan-queue-empty');
        const toolLoanSubmitLabel = document.getElementById('tool-loan-submit-label');
        const toolLoansListFilterBorrower = document.getElementById('tool-loans-list-filter-borrower');
        const toolLoansListFilterProduct = document.getElementById('tool-loans-list-filter-product');
        const toolLoansFilterReturnedStatus = document.getElementById('tool-loans-filter-returned-status');
        const importLocationsBtn = document.getElementById('import-locations-btn');
        const csvLocationsFileInput = document.getElementById('csv-locations-file-input');
        const feedbackFooterBtn = document.getElementById('feedback-footer-btn');
        const feedbackBackdrop = document.getElementById('feedback-backdrop');
        const feedbackPanel = document.getElementById('feedback-panel');
        const feedbackForm = document.getElementById('feedback-form');
        const feedbackMessage = document.getElementById('feedback-message');
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


        // --- Utilitários de Autenticação ---
        const hashPassword = async (password) => {
            const encoder = new TextEncoder();
            const data = encoder.encode(password);
            const hashBuffer = await crypto.subtle.digest('SHA-256', data);
            const hashArray = Array.from(new Uint8Array(hashBuffer));
            const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
            return hashHex;
        };

        const normalizeUserId = (username) =>
            username.toLowerCase().trim().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');

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
                
                if (userDoc.exists) {
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

        /** Confirmação de exclusão no estoque com PIN anti-clique acidental. */
        const showStockDeleteConfirmationModal = (title, message, onConfirm) => {
            const content = `
                <h2 class="text-2xl font-bold mb-4">${title}</h2>
                <p class="text-slate-600 mb-4">${message}</p>
                <p class="text-sm text-slate-500 mb-2">Digite o PIN para confirmar.</p>
                <input type="password" inputmode="numeric" pattern="[0-9]*" autocomplete="off" id="stock-delete-pin-input" maxlength="8" class="w-full p-3 border border-slate-200 rounded-lg text-lg tracking-widest text-center font-mono mb-6 focus:ring-2 focus:ring-red-500/40 focus:border-red-500" placeholder="PIN">
                <div class="flex justify-end gap-4">
                    <button type="button" id="stock-delete-cancel" class="px-6 py-2 bg-slate-200 text-slate-800 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                    <button type="button" id="stock-delete-ok" class="px-6 py-2 bg-red-600 text-white rounded-lg font-semibold hover:bg-red-700 transition">Excluir</button>
                </div>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            openModal('generic-modal');
            const pinInput = document.getElementById('stock-delete-pin-input');
            pinInput?.focus();
            document.getElementById('stock-delete-ok').onclick = () => {
                const pin = (pinInput?.value || '').trim();
                if (pin !== STOCK_DELETE_PIN) {
                    showToast('PIN incorreto.', true);
                    return;
                }
                onConfirm();
                closeModal('generic-modal');
            };
            document.getElementById('stock-delete-cancel').onclick = () => closeModal('generic-modal');
        };

        /** Canvas de assinatura para tablet/desktop; retorna { clear, hasInk, toDataURL }. */
        const setupSignaturePad = (canvas, clearButton) => {
            if (!canvas) return null;
            if (canvas.dataset.padWired === '1') return null;
            const ctx = canvas.getContext('2d');
            const W = 560;
            const H = 160;
            canvas.width = W;
            canvas.height = H;
            canvas.style.width = '100%';
            canvas.style.maxWidth = '560px';
            canvas.style.height = '160px';
            canvas.style.touchAction = 'none';
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, W, H);
            ctx.strokeStyle = '#0f172a';
            ctx.lineWidth = 2.25;
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';
            canvas.dataset.hasInk = '0';

            let drawing = false;
            const getPoint = (e) => {
                const r = canvas.getBoundingClientRect();
                const t = e.touches && e.touches[0];
                const clientX = t ? t.clientX : e.clientX;
                const clientY = t ? t.clientY : e.clientY;
                return {
                    x: ((clientX - r.left) / r.width) * canvas.width,
                    y: ((clientY - r.top) / r.height) * canvas.height
                };
            };
            const start = (e) => {
                if (e.cancelable) e.preventDefault();
                drawing = true;
                const { x, y } = getPoint(e);
                ctx.beginPath();
                ctx.moveTo(x, y);
            };
            const move = (e) => {
                if (!drawing) return;
                if (e.cancelable) e.preventDefault();
                const { x, y } = getPoint(e);
                ctx.lineTo(x, y);
                ctx.stroke();
                ctx.beginPath();
                ctx.moveTo(x, y);
                canvas.dataset.hasInk = '1';
            };
            const end = () => { drawing = false; };

            canvas.addEventListener('mousedown', start);
            canvas.addEventListener('mousemove', move);
            canvas.addEventListener('mouseup', end);
            canvas.addEventListener('mouseleave', end);
            canvas.addEventListener('touchstart', start, { passive: false });
            canvas.addEventListener('touchmove', move, { passive: false });
            canvas.addEventListener('touchend', end);
            canvas.addEventListener('touchcancel', end);

            const clear = () => {
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                canvas.dataset.hasInk = '0';
            };
            clearButton?.addEventListener('click', (ev) => {
                ev.preventDefault();
                clear();
            });

            canvas.dataset.padWired = '1';

            return {
                clear,
                hasInk: () => canvas.dataset.hasInk === '1',
                toDataURL: () => canvas.toDataURL('image/jpeg', 0.85)
            };
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
                logoContainer.innerHTML = `<span class="material-symbols-outlined text-white" style="font-size:22px;">inventory_2</span>`;
                logoContainer.className = 'p-2.5 rounded-xl flex items-center justify-center h-11 w-11';
                logoContainer.style.background = 'linear-gradient(135deg, #0066FF, #4F46E5)';
                logoContainer.style.boxShadow = '0 4px 12px rgba(0,102,255,0.3)';
            }
        };

        const initializeAppSession = async (customUser) => {
            currentUser = customUser;
            userRole = customUser.role;
            currentObraId = customUser.obraId || 'uhe_estrela';
            if (loginBtn) loginBtn.disabled = false;

            // Determinar base de dados - Simplificado para priorizar a base principal
            const obraId = customUser.obraId || 'uhe_estrela';
            const obraBase = obraId === 'uhe_estrela'
                ? `/artifacts/${appId}/public/data`
                : `/artifacts/${appId}/public/data/obras/${obraId}`;

            // Configurar referências do Firestore
            productsCollectionRef = collection(db, `${obraBase}/products`);
            historyCollectionRef = collection(db, `${obraBase}/history`);
            requisitionsCollectionRef = collection(db, `${obraBase}/requisitions`);
            toolLoansCollectionRef = collection(db, `${obraBase}/tool_loans`);
            locationsCollectionRef = collection(db, `${obraBase}/locations`);
            settingsDocRef = doc(db, `${obraBase}/app_settings/main`);
            usersCollectionRef = collection(db, `/artifacts/${appId}/public/data/users`);

            // ⭐ UHE Estrela collections (exclusivo para obra uhe_estrela)
            estrelaProductsRef = collection(db, `/artifacts/${appId}/public/data/estrela_products`);
            estrelaEntriesRef = collection(db, `/artifacts/${appId}/public/data/estrela_entries`);
            estrelaExitsRef = collection(db, `/artifacts/${appId}/public/data/estrela_exits`);

            // Mostrar/ocultar aba Estoque Estrela conforme a obra
            const navSecaoEstrela = document.getElementById('nav-section-estrela');
            const navBtnEstrela = document.getElementById('nav-btn-estrela');
            if (obraId !== 'uhe_estrela') {
                if (navSecaoEstrela) navSecaoEstrela.style.display = 'none';
                if (navBtnEstrela) navBtnEstrela.style.display = 'none';
            } else {
                if (navSecaoEstrela) navSecaoEstrela.style.display = '';
                if (navBtnEstrela) navBtnEstrela.style.display = '';
            }

            // Atualizar nome do app no header e sidebar conforme a obra
            const obraDisplayNames = { 'uhe_estrela': 'UHE Estrela', 'pch_taboca': 'PCH Taboca' };
            const obraDisplayName = obraDisplayNames[obraId] || customUser.displayName;
            const appTitleEl = document.getElementById('app-title');
            const sidebarAppTitleEl = document.getElementById('sidebar-app-title');
            if (appTitleEl) appTitleEl.textContent = obraDisplayName;
            if (sidebarAppTitleEl) sidebarAppTitleEl.textContent = obraDisplayName;

            const initials = customUser.displayName ? customUser.displayName.split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase() : '??';
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
            
            // 🔒 Timeout de segurança para garantir carregamento
            setTimeout(() => {
                if (!isDataLoaded) {
                    isDataLoaded = true;
                    showLoader(false);
                    showProductsSkeleton(false);
                    console.warn('⚠️ Carregamento com timeout: dados podem estar vazios');
                }
            }, 5000);
            
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
            showLoader(false); // 🚀 Forçar saída do loader em caso de erro

            let errorMessage = '';
            switch (error.code) {
                case 'permission-denied':
                    errorMessage = `🔒 Sem permissão para acessar ${dataType}.`;
                    break;
                case 'unavailable':
                    errorMessage = `📡 Servidor indisponível. Verifique a conexão.`;
                    break;
                case 'not-found':
                    errorMessage = `❓ Coleção de ${dataType} não encontrada.`;
                    break;
                default:
                    errorMessage = `❌ Erro ao carregar ${dataType}. Tente recarregar a página.`;
            }

            showToast(errorMessage, true);
        };

        function stopSnapshotGroup(unsubscribers) {
            unsubscribers.forEach(unsub => {
                try { if (typeof unsub === 'function') unsub(); } catch (_) {}
            });
            unsubscribers.length = 0;
        }

        function stopCoreListeners() {
            stopSnapshotGroup(coreUnsubscribers);
        }

        function stopEstrelaListeners() {
            stopSnapshotGroup(estrelaUnsubscribers);
        }

        function startCoreListeners() {
            if (coreUnsubscribers.length > 0) return;

            const obraDefaultNames = { 'uhe_estrela': 'UHE Estrela', 'pch_taboca': 'PCH Taboca' };
            const obraDefaultName = obraDefaultNames[currentUser?.obraId] || 'UHE Estrela';
            coreUnsubscribers.push(onSnapshot(settingsDocRef, (doc) => {
                if (doc.exists) {
                    const data = doc.data();
                    // Corrigir nomes legados apenas para UHE Estrela
                    if (currentUser?.obraId === 'uhe_estrela' && (data.appName === 'Estoque Taboca' || data.appName === 'Estoque Estrela')) {
                        data.appName = 'UHE Estrela';
                        setDoc(settingsDocRef, { appName: 'UHE Estrela' }, { merge: true }).catch(() => {});
                    }
                    updateAppSettingsUI(data);
                }
                else updateAppSettingsUI({ appName: obraDefaultName, logoUrl: null });
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
                    renderToolLoans();
                }
            }, (error) => handleFirestoreError(error, 'produtos')));

            coreUnsubscribers.push(onSnapshot(historyCollectionRef, (snapshot) => {
                history = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (isDataLoaded) {
                    updateDashboard();
                    renderExitLog(exitsSearchInput.value);
                    renderActivityLog(activityLogSearchInput?.value || '');
                    renderRMView(rmSearchInput.value);
                    tryOpenPlaqueDeepLink();
                }
            }, (error) => handleFirestoreError(error, 'histórico')));

            coreUnsubscribers.push(onSnapshot(requisitionsCollectionRef, (snapshot) => {
                requisitions = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (isDataLoaded) renderRequisitions();
            }, (error) => handleFirestoreError(error, 'requisições')));

            coreUnsubscribers.push(onSnapshot(toolLoansCollectionRef, (snapshot) => {
                toolLoans = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                if (isDataLoaded) {
                    renderToolLoans();
                    if (currentViewId === 'dashboard-view') updateDashboard();
                    // Atualiza badge de cautelas em aberto em qualquer view
                    const openCount = toolLoans.filter(l => l.status !== 'returned' && l.status !== 'damaged').length;
                    [document.getElementById('nav-tool-loans-badge'), document.getElementById('drawer-tool-loans-badge')].forEach(badge => {
                        if (!badge) return;
                        if (openCount > 0) { badge.textContent = openCount > 99 ? '99+' : openCount; badge.classList.remove('hidden'); }
                        else { badge.classList.add('hidden'); }
                    });
                }
            }, (error) => handleFirestoreError(error, 'cautelas')));

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
                    renderToolLoans();
                    setTimeout(() => tryOpenPlaqueDeepLink(), 1200);
                } else {
                    renderLocations();
                }
            }, (error) => {
                console.error('Erro ao buscar locais:', error);
                if (!isDataLoaded) {
                    isDataLoaded = true;
                    showLoader(false);
                    showProductsSkeleton(false);
                }
                handleFirestoreError(error, 'locais');
            }));
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

        /**
         * Mantém listeners ativos enquanto o usuário está logado (mesmo com a aba em segundo plano).
         * Antes: ao ocultar a aba, todos os listeners eram removidos e, ao voltar, o Firestore cobrava
         * de novo uma leitura por documento — muito caro em histórico grande. Logout ainda encerra tudo.
         */
        function syncRealtimeListeners() {
            if (!currentUser) {
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
                    productName: toUpperText(product?.name || 'Produto Excluído'),
                    type, quantity: Math.abs(quantity), newTotal,
                    withdrawnBy: toUpperOrNull(details.withdrawnBy),
                    teamLeader: toUpperOrNull(details.teamLeader),
                    applicationLocation: toUpperOrNull(details.applicationLocation),
                    obra: details.obra || null,
                    details: details.details || null,
                    rmProcessed: false,
                    date: serverTimestamp(),
                    performedBy: toUpperText(currentUser?.displayName || 'Anônimo'), // 🔍 Auditoria: usuário que realizou ação
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

            // Cautelas abertas do dia
            const todayStr = new Date().toLocaleDateString('pt-BR');
            const todayLoans = toolLoans.filter(loan => {
                if (!loan.loanDate) return false;
                const ms = typeof loan.loanDate.seconds === 'number'
                    ? loan.loanDate.seconds * 1000
                    : (typeof loan.loanDate.toMillis === 'function' ? loan.loanDate.toMillis() : 0);
                return ms > 0 && new Date(ms).toLocaleDateString('pt-BR') === todayStr;
            });
            const renderTodayLoans = () => {
                if (todayLoans.length === 0) {
                    return `<p class="text-xs mt-1" style="color:#727785;">Nenhuma cautela registrada hoje.</p>`;
                }
                return `<div class="mt-2 space-y-1.5">
                    ${todayLoans.map(l => `
                        <div class="flex items-start gap-2 px-2 py-1.5 rounded-lg" style="background:#f8f9fa;">
                            <span class="material-symbols-outlined shrink-0 mt-0.5" style="font-size:15px; color:#795900; font-variation-settings:'FILL' 1,'wght' 400;">construction</span>
                            <div class="min-w-0">
                                <p class="text-xs font-bold truncate" style="color:#191c1d;">${l.borrower || '—'}</p>
                                <p class="text-xs truncate" style="color:#727785;">${l.productName || '—'}</p>
                            </div>
                        </div>`).join('')}
                </div>`;
            };

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

            // Atualizar badge de cautelas em aberto (sidebar + drawer)
            const openToolLoansCount = toolLoans.filter(l => l.status !== 'returned' && l.status !== 'damaged').length;
            [document.getElementById('nav-tool-loans-badge'), document.getElementById('drawer-tool-loans-badge')].forEach(badge => {
                if (!badge) return;
                if (openToolLoansCount > 0) {
                    badge.textContent = openToolLoansCount > 99 ? '99+' : openToolLoansCount;
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
                <!-- KPI: Cautelas do Dia -->
                <div class="bg-surface-container-lowest p-5 rounded-2xl tonal-elevation dashboard-reveal" style="animation-delay: 0.2s; border-bottom: 3px solid #d97706;">
                    <div class="flex items-center justify-between mb-2">
                        <div class="p-2 rounded-xl" style="background:rgba(121,89,0,0.1);">
                            <span class="material-symbols-outlined" style="font-size:20px; color:#795900; font-variation-settings:'FILL' 1,'wght' 400;">construction</span>
                        </div>
                        <span class="text-xs font-extrabold px-2 py-0.5 rounded-full" style="background:rgba(217,119,6,0.12); color:#92400e;">${todayLoans.length}</span>
                    </div>
                    <p class="text-xs font-bold uppercase tracking-wide" style="color:#795900;">Cautelas do Dia</p>
                    ${renderTodayLoans()}
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
                    <button onclick="document.querySelector('[data-view=activity-log-view]').click()" class="text-xs font-bold uppercase tracking-widest" style="color:#005bbf; letter-spacing:0.08em;">Ver tudo</button>
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

        const toUpperText = (value = '') => String(value ?? '').trim().toLocaleUpperCase('pt-BR');

        /** Só http/https; evita javascript: e lixo no src de <img>. */
        const sanitizeProductImageUrl = (raw) => {
            const t = String(raw ?? '').trim();
            if (!t) return null;
            try {
                const u = new URL(t);
                if (u.protocol !== 'http:' && u.protocol !== 'https:') return null;
                return u.toString();
            } catch {
                return null;
            }
        };

        const escapeHtmlAttr = (s) => String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');

        const updateProductPhotoPreview = (inputId, wrapId, imgId) => {
            const inp = document.getElementById(inputId);
            const wrap = document.getElementById(wrapId);
            const img = document.getElementById(imgId);
            if (!inp || !wrap || !img) return;
            const u = sanitizeProductImageUrl(inp.value);
            if (!u) {
                wrap.classList.add('hidden');
                img.removeAttribute('src');
                return;
            }
            img.onload = () => { wrap.classList.remove('hidden'); };
            img.onerror = () => {
                img.removeAttribute('src');
                wrap.classList.add('hidden');
            };
            img.referrerPolicy = 'no-referrer';
            img.decoding = 'async';
            img.src = u;
        };

        const openProductImageLightbox = (url) => {
            const safe = sanitizeProductImageUrl(url);
            if (!safe) return;
            const root = document.getElementById('product-image-lightbox');
            const img = document.getElementById('product-image-lightbox-img');
            if (!root || !img) return;
            img.referrerPolicy = 'no-referrer';
            img.src = safe;
            root.classList.remove('hidden');
            root.setAttribute('aria-hidden', 'false');
            document.body.style.overflow = 'hidden';
        };

        const closeProductImageLightbox = () => {
            const root = document.getElementById('product-image-lightbox');
            const img = document.getElementById('product-image-lightbox-img');
            if (img) img.removeAttribute('src');
            root?.classList.add('hidden');
            root?.setAttribute('aria-hidden', 'true');
            document.body.style.overflow = '';
        };
        const toUpperOrNull = (value = '') => {
            const normalized = toUpperText(value);
            return normalized || null;
        };
        const parseUpperAliases = (value = '') => String(value || '')
            .split(',')
            .map(alias => toUpperText(alias))
            .filter(Boolean);

        const uppercaseRecordFields = (record = {}, fields = []) => {
            const normalized = { ...record };
            fields.forEach(field => {
                if (normalized[field] !== undefined && normalized[field] !== null) {
                    normalized[field] = toUpperText(normalized[field]);
                }
            });
            return normalized;
        };

        const normalizeRecordForUppercase = (collectionName, record = {}) => {
            if (collectionName === 'products' || collectionName === 'estrela_products') {
                return uppercaseRecordFields(record, ['name', 'location', 'createdBy', 'updatedBy']);
            }
            if (collectionName === 'locations') {
                return {
                    ...uppercaseRecordFields(record, ['name']),
                    aliases: Array.isArray(record.aliases) ? record.aliases.map(alias => toUpperText(alias)).filter(Boolean) : record.aliases
                };
            }
            if (collectionName === 'requisitions') {
                return {
                    ...uppercaseRecordFields(record, ['requester', 'teamLeader', 'applicationLocation']),
                    items: Array.isArray(record.items)
                        ? record.items.map(item => uppercaseRecordFields(item, ['productName']))
                        : record.items
                };
            }
            if (collectionName === 'toolLoans' || collectionName === 'tool_loans') {
                return uppercaseRecordFields(record, ['productName', 'borrower', 'role', 'observation', 'createdBy', 'returnedBy']);
            }
            if (collectionName === 'history') {
                return uppercaseRecordFields(record, ['productName', 'withdrawnBy', 'teamLeader', 'applicationLocation', 'performedBy', 'receivedBy', 'supplier', 'observation']);
            }
            return record;
        };

        const uppercaseInputSelector = [
            '#product-name',
            '#product-location',
            '#edit-product-name',
            '#edit-product-location',
            '#location-name',
            '#location-aliases',
            '#edit-location-name',
            '#edit-location-aliases',
            '#entry-supplier',
            '#entry-observation',
            '#req-requester',
            '#req-team-leader',
            '#req-application-location',
            '#tool-loan-search',
            '#tool-loan-borrower',
            '#tool-loan-role',
            '#tool-loan-observation',
            '#ep-name',
            '#ep-location',
            '#epe-name',
            '#epe-location',
            '#estrela-entry-supplier',
            '#estrela-entry-received-by',
            '#estrela-entry-obs',
            '#estrela-exit-who',
            '#estrela-exit-leader',
            '#estrela-exit-location',
            '#estrela-exit-obs'
        ].join(',');

        document.body.addEventListener('input', (e) => {
            if (!e.target.matches?.(uppercaseInputSelector)) return;
            const start = e.target.selectionStart;
            const end = e.target.selectionEnd;
            e.target.value = e.target.value.toLocaleUpperCase('pt-BR');
            if (typeof e.target.setSelectionRange === 'function' && start !== null && end !== null) {
                e.target.setSelectionRange(start, end);
            }
        });

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
        const populateLocationFilter = () => {
            if (!locationFilterSelect) return;
            const uniqueLocations = [...new Set(
                products
                    .map(p => (p.location || '').trim())
                    .filter(loc => loc.length > 0)
            )].sort((a, b) => a.localeCompare(b, 'pt-BR'));

            const previousValue = inventoryLocationFilter;
            const stillExists = previousValue === 'all' || uniqueLocations.includes(previousValue);

            locationFilterSelect.innerHTML = '<option value="all">Todas as localizações</option>' +
                uniqueLocations.map(loc => `<option value="${loc.replace(/"/g, '&quot;')}">${loc}</option>`).join('');

            if (stillExists) {
                locationFilterSelect.value = previousValue;
            } else {
                inventoryLocationFilter = 'all';
                locationFilterSelect.value = 'all';
            }
        };

        const renderProducts = () => {
            populateLocationFilter();
            const filterText = searchInput.value.toLowerCase();
            let processedProducts = [...products]; 
            if (inventoryFilter === 'low_stock') {
                processedProducts = processedProducts.filter(p => p.quantity <= p.minQuantity);
            }
            if (inventoryLocationFilter && inventoryLocationFilter !== 'all') {
                processedProducts = processedProducts.filter(p => (p.location || '').trim() === inventoryLocationFilter);
            }
            if (filterText) {
                processedProducts = processedProducts.filter(p =>
                    (p.name && p.name.toLowerCase().includes(filterText)) || 
                    (p.code && p.code.toLowerCase().includes(filterText)) ||
                    (p.codeRM && p.codeRM.toLowerCase().includes(filterText)) ||
                    (p.location && p.location.toLowerCase().includes(filterText))
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
                const safeImg = sanitizeProductImageUrl(p.imageUrl);
                const thumbBlock = safeImg
                    ? `<button type="button" class="product-thumb-lightbox shrink-0 w-16 h-16 sm:w-20 sm:h-20 rounded-lg border border-slate-200 bg-slate-100 overflow-hidden relative flex items-center justify-center p-0 cursor-zoom-in hover:ring-2 hover:ring-indigo-400/60 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition" data-image-url="${escapeHtmlAttr(safeImg)}" title="Ver foto maior" aria-label="Ampliar foto do produto">
                        <img src="${escapeHtmlAttr(safeImg)}" alt="" width="80" height="80" class="pointer-events-none w-full h-full object-cover" loading="lazy" decoding="async" referrerpolicy="no-referrer" onerror="this.classList.add('hidden');this.nextElementSibling.classList.remove('hidden');">
                        <span class="material-symbols-outlined pointer-events-none text-slate-300 text-4xl hidden" style="font-variation-settings:'FILL' 0;">broken_image</span>
                       </button>`
                    : `<div class="shrink-0 w-16 h-16 sm:w-20 sm:h-20 rounded-lg border border-dashed border-slate-200 bg-slate-50 flex items-center justify-center text-slate-300" title="Sem foto">
                        <span class="material-symbols-outlined text-3xl">image</span>
                       </div>`;
                const tr = document.createElement('tr');
                tr.className = `hover:bg-slate-50 transition-colors duration-150`;
                tr.innerHTML = `
                    <td class="p-3 sm:p-4 text-center"><input type="checkbox" class="product-checkbox h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-600" data-id="${p.id}" ${isChecked ? 'checked' : ''}></td>
                    <td class="p-3 sm:p-4 align-top">
                        <div class="flex gap-3 items-start">
                            ${thumbBlock}
                            <div class="min-w-0 flex-1">
                                <p class="font-bold text-slate-800 text-sm sm:text-base leading-tight">${p.name}</p>
                                <p class="text-xs text-slate-500 mt-0.5">RM: <span class="font-medium">${p.codeRM || 'N/A'}</span></p>
                                <p class="text-xs text-slate-400">SKU: ${p.code}</p>
                                <p class="text-xs text-slate-500 sm:hidden mt-0.5">${p.location || ''}</p>
                            </div>
                        </div>
                    </td>
                    <td class="p-3 sm:p-4 align-top text-slate-600 text-sm hidden md:table-cell">${p.group || 'N/A'}</td>
                    <td class="p-3 sm:p-4 align-top text-slate-600 text-sm hidden md:table-cell">${p.unit || 'N/A'}</td>
                    <td class="p-3 sm:p-4 align-top"><div class="flex flex-col sm:flex-row sm:items-center gap-1"><p class="text-base sm:text-lg font-bold text-slate-800">${p.quantity}</p>${isLowStock ? '<span class="px-1.5 py-0.5 text-xs font-semibold text-yellow-800 bg-yellow-100 rounded-full">Baixo</span>' : ''}</div></td>
                    <td class="p-3 sm:p-4 align-top text-slate-600 text-sm hidden sm:table-cell">${p.minQuantity}</td>
                    <td class="p-3 sm:p-4 align-top text-slate-600 text-sm hidden sm:table-cell">${p.location}</td>
                    <td class="p-3 sm:p-4 align-top text-center"><div class="flex justify-center items-center gap-0.5 sm:gap-1"><button data-id="${p.id}" class="history-btn text-slate-500 hover:text-purple-600 p-1.5 sm:p-2 rounded-full hover:bg-purple-100 transition" title="Histórico"><svg class="w-4 h-4 sm:w-5 sm:h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg></button><button data-id="${p.id}" class="edit-btn text-slate-500 hover:text-blue-600 p-1.5 sm:p-2 rounded-full hover:bg-blue-100 transition" title="Editar"><svg class="w-4 h-4 sm:w-5 sm:h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg></button><button data-id="${p.id}" class="delete-btn text-slate-500 hover:text-red-600 p-1.5 sm:p-2 rounded-full hover:bg-red-100 transition" title="Excluir"><svg class="w-4 h-4 sm:w-5 sm:h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg></button></div></td>`;
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

        const formatFirestoreDate = (value) => {
            if (!value) return '...';
            if (typeof value.seconds === 'number') return new Date(value.seconds * 1000).toLocaleString('pt-BR');
            if (typeof value.toMillis === 'function') return new Date(value.toMillis()).toLocaleString('pt-BR');
            return '...';
        };

        const populateToolLoanProducts = () => {
            if (!toolLoanProductSelect) return;
            const currentValue = toolLoanProductSelect.value;
            const searchTerm = normalizeSearchText(toolLoanSearchInput?.value || '');
            const tokens = searchTerm.split(/\s+/).filter(Boolean);
            const filteredProducts = products.filter(product => {
                if (tokens.length === 0) return true;
                const haystack = normalizeSearchText([product.name, product.codeRM, product.code, product.location, product.group].filter(Boolean).join(' '));
                return tokens.every(token => haystack.includes(token));
            });
            const sortedProducts = filteredProducts.sort((a, b) => (a.name || '').localeCompare(b.name || ''));

            toolLoanProductSelect.innerHTML = `<option value="">${sortedProducts.length ? 'Selecione a ferramenta' : 'Nenhuma ferramenta encontrada'}</option>` + sortedProducts
                .map(p => {
                    const reserved = toolLoanQueue.filter(l => l.productId === p.id).reduce((s, l) => s + l.quantity, 0);
                    const disp = Math.max(0, (p.quantity || 0) - reserved);
                    return `<option value="${p.id}">${p.name} (${p.codeRM || p.code || 'S/C'}) — Disp.: ${disp}</option>`;
                })
                .join('');

            if (currentValue && sortedProducts.some(p => p.id === currentValue)) {
                toolLoanProductSelect.value = currentValue;
            }
        };

        const getToolLoansListFilterTokens = (el) => {
            const raw = el?.value || '';
            return normalizeSearchText(raw).split(/\s+/).filter(Boolean);
        };

        const loanMatchesListFilters = (loan, borrowerTokens, productTokens) => {
            if (borrowerTokens.length) {
                const bh = normalizeSearchText([loan.borrower, loan.role].filter(Boolean).join(' '));
                if (!borrowerTokens.every(t => bh.includes(t))) return false;
            }
            if (productTokens.length) {
                const ph = normalizeSearchText([loan.productName, loan.productCodeRM, loan.productCode].filter(Boolean).join(' '));
                if (!productTokens.every(t => ph.includes(t))) return false;
            }
            return true;
        };

        const getToolLoanQueueQtyForProduct = (productId) =>
            toolLoanQueue.filter(l => l.productId === productId).reduce((s, l) => s + l.quantity, 0);

        const renderToolLoanQueue = () => {
            if (!toolLoanQueueBody) return;
            toolLoanQueueBody.innerHTML = '';
            toolLoanQueue.forEach((line, idx) => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="p-2 pl-3 font-medium text-slate-800">${line.productName || ''}</td>
                    <td class="p-2 text-slate-600 text-xs">${line.productCodeRM || line.productCode || '—'}</td>
                    <td class="p-2 text-center font-bold text-slate-800">${line.quantity}</td>
                    <td class="p-2 text-center">
                        <button type="button" data-queue-idx="${idx}" class="tool-loan-queue-remove text-red-600 hover:bg-red-50 p-1 rounded" title="Remover da lista">
                            <span class="material-symbols-outlined align-middle" style="font-size:20px;">close</span>
                        </button>
                    </td>`;
                toolLoanQueueBody.appendChild(tr);
            });
            if (toolLoanQueueEmpty) {
                toolLoanQueueEmpty.classList.toggle('hidden', toolLoanQueue.length > 0);
            }
            if (toolLoanSubmitLabel) {
                if (toolLoanQueue.length > 1) {
                    toolLoanSubmitLabel.textContent = `Cautelar ${toolLoanQueue.length} itens`;
                } else if (toolLoanQueue.length === 1) {
                    toolLoanSubmitLabel.textContent = 'Cautelar 1 item';
                } else {
                    toolLoanSubmitLabel.textContent = 'Cautelar';
                }
            }
            populateToolLoanProducts();
        };

        const renderToolLoans = () => {
            populateToolLoanProducts();
            if (!toolLoansOpenList || !toolLoansReturnedList) return;

            const borrowerTokens = getToolLoansListFilterTokens(toolLoansListFilterBorrower);
            const productTokens = getToolLoansListFilterTokens(toolLoansListFilterProduct);
            const returnedStatus = toolLoansFilterReturnedStatus?.value || 'all';
            const hasListFilters = borrowerTokens.length > 0 || productTokens.length > 0 || returnedStatus !== 'all';

            const sortedLoans = [...toolLoans].sort((a, b) => (b.loanDate?.seconds || 0) - (a.loanDate?.seconds || 0));
            let openLoans = sortedLoans.filter(loan => loan.status !== 'returned' && loan.status !== 'damaged');
            let returnedLoans = sortedLoans.filter(loan => loan.status === 'returned' || loan.status === 'damaged');

            openLoans = openLoans.filter(loan => loanMatchesListFilters(loan, borrowerTokens, productTokens));
            returnedLoans = returnedLoans.filter(loan => loanMatchesListFilters(loan, borrowerTokens, productTokens));
            if (returnedStatus === 'returned') {
                returnedLoans = returnedLoans.filter(l => l.status === 'returned');
            } else if (returnedStatus === 'damaged') {
                returnedLoans = returnedLoans.filter(l => l.status === 'damaged');
            }

            toolLoansOpenList.innerHTML = '';
            toolLoansReturnedList.innerHTML = '';

            noToolLoansOpenMessage?.classList.toggle('hidden', openLoans.length > 0);
            noToolLoansReturnedMessage?.classList.toggle('hidden', returnedLoans.length > 0);
            if (openLoans.length === 0 && isDataLoaded && noToolLoansOpenMessage) {
                noToolLoansOpenMessage.innerHTML = hasListFilters
                    ? '<p class="text-lg">Nenhuma cautela em aberto com esses filtros.</p>'
                    : '<p class="text-lg">Nenhuma ferramenta em aberto.</p>';
            }
            if (returnedLoans.length === 0 && isDataLoaded && noToolLoansReturnedMessage) {
                noToolLoansReturnedMessage.innerHTML = hasListFilters
                    ? '<p class="text-lg">Nenhum registro na coluna devolvidas com esses filtros.</p>'
                    : '<p class="text-lg">Nenhuma devolução registrada.</p>';
            }

            openLoans.forEach(loan => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                tr.innerHTML = `
                    <td class="p-3 text-sm text-slate-600">${formatFirestoreDate(loan.loanDate)}</td>
                    <td class="p-3"><p class="font-semibold text-slate-800">${loan.productName || ''}</p><p class="text-sm text-slate-500">${loan.productCodeRM || loan.productCode || 'N/A'}</p></td>
                    <td class="p-3 text-center font-bold text-slate-800">${loan.quantity || 0}</td>
                    <td class="p-3"><p class="font-semibold text-slate-700">${loan.borrower || ''}</p><p class="text-sm text-slate-500">${loan.role || ''}</p></td>
                    <td class="p-3 text-center">
                        <button data-id="${loan.id}" class="return-tool-loan-btn bg-green-600 text-white text-xs font-bold px-3 py-2 rounded-lg hover:bg-green-700 transition">Devolvido</button>
                        <button data-id="${loan.id}" class="damaged-tool-loan-btn bg-red-600 text-white text-xs font-bold px-3 py-2 rounded-lg hover:bg-red-700 transition ml-1">D. Danificado</button>
                    </td>
                `;
                toolLoansOpenList.appendChild(tr);
            });

            returnedLoans.slice(0, 100).forEach(loan => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                const isDamaged = loan.status === 'damaged';
                const badge = isDamaged
                    ? `<span class="ml-1 inline-block px-1.5 py-0.5 rounded text-xs font-bold" style="background:#fee2e2;color:#b91c1c;">DANIFICADA</span>`
                    : `<span class="ml-1 inline-block px-1.5 py-0.5 rounded text-xs font-bold" style="background:#dcfce7;color:#15803d;">DEVOLVIDA</span>`;
                tr.innerHTML = `
                    <td class="p-3 text-sm text-slate-600">${formatFirestoreDate(loan.returnDate)}</td>
                    <td class="p-3"><p class="font-semibold text-slate-800">${loan.productName || ''}${badge}</p><p class="text-sm text-slate-500">${loan.productCodeRM || loan.productCode || 'N/A'}</p></td>
                    <td class="p-3 text-center font-bold text-slate-800">${loan.quantity || 0}</td>
                    <td class="p-3 text-slate-700">${loan.borrower || ''}</td>
                `;
                toolLoansReturnedList.appendChild(tr);
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

        const escHtmlText = (s) => String(s ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');

        const buildActivityLogHaystack = (h) => [h.type, h.productName, h.productCode, h.productCodeRM, h.withdrawnBy, h.performedBy, h.receivedBy, h.supplier, h.nfNumber, h.applicationLocation, h.teamLeader, h.details, h.obra]
            .filter((v) => v != null && String(v).trim() !== '')
            .map((v) => String(v).toLowerCase())
            .join(' ');

        const buildActivityLogDetail = (h) => {
            const parts = [];
            const actor = h.performedBy || h.receivedBy || h.withdrawnBy;
            if (actor) parts.push(String(actor));
            if (h.details) parts.push(String(h.details));
            if (h.nfNumber) parts.push(`NF ${h.nfNumber}`);
            if (h.supplier) parts.push(`Fornec.: ${h.supplier}`);
            if (h.applicationLocation) parts.push(`Aplicação: ${h.applicationLocation}`);
            if (h.teamLeader) parts.push(`Função: ${h.teamLeader}`);
            if (h.obra) parts.push(`Obra: ${h.obra}`);
            return parts.length ? parts.join(' · ') : '—';
        };

        const formatActivityLogQuantity = (h) => {
            const t = h.type || '';
            if (t === 'Edição') return '<span class="text-slate-400 font-medium">—</span>';
            const q = Math.abs(Number(h.quantity) || 0);
            const positiveTypes = ['Entrada', 'Entrada por NF', 'Ajuste Entrada', 'Criação', 'Importação'];
            const isPositive = positiveTypes.includes(t);
            const cls = isPositive ? 'text-green-600' : 'text-red-600';
            const sym = isPositive ? '+' : '−';
            return `<span class="font-bold text-lg ${cls}">${sym}${q}</span>`;
        };

        const renderActivityLog = (filter = '') => {
            if (!activityLogList || !noActivityLogMessage) return;
            const q = (filter || '').toLowerCase().trim();
            const filtered = [...history].filter((h) => {
                if (!q) return true;
                return buildActivityLogHaystack(h).includes(q);
            }).sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            activityLogList.innerHTML = '';
            noActivityLogMessage.classList.toggle('hidden', !(filtered.length === 0 && isDataLoaded));
            if (filtered.length === 0 && isDataLoaded) {
                noActivityLogMessage.innerHTML = q
                    ? '<p class="text-lg">Nenhum registro encontrado para esta busca.</p>'
                    : '<p class="text-lg">Nenhuma atividade registrada ainda.</p>';
            }

            filtered.forEach((h) => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition-colors duration-150';
                const dateStr = h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : '…';
                const saldo = h.newTotal != null && h.newTotal !== '' ? String(h.newTotal) : '—';
                const detail = escHtmlText(buildActivityLogDetail(h));
                tr.innerHTML = `
                    <td class="p-3 sm:p-4 text-sm text-slate-600 whitespace-nowrap">${dateStr}</td>
                    <td class="p-3 sm:p-4 text-sm font-semibold text-slate-800">${escHtmlText(h.type || '—')}</td>
                    <td class="p-3 sm:p-4"><p class="font-semibold text-slate-800 text-sm">${escHtmlText(h.productName)}</p><p class="text-xs text-slate-500">${escHtmlText(h.productCodeRM || h.productCode || '')}</p></td>
                    <td class="p-3 sm:p-4 text-center">${formatActivityLogQuantity(h)}</td>
                    <td class="p-3 sm:p-4 text-center text-slate-700 font-semibold hidden sm:table-cell">${escHtmlText(saldo)}</td>
                    <td class="p-3 sm:p-4 text-sm text-slate-600 max-w-md">${detail}</td>
                `;
                activityLogList.appendChild(tr);
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
                        <div><label class="block text-sm font-medium">Função do Funcionário</label><input type="text" id="req-team-leader" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required></div>
                        <div class="md:col-span-2">
                            <label class="block text-sm font-medium">Local de Aplicação</label>
                            <input type="text" id="req-application-location" name="req-application-location" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" placeholder="Digite o local de aplicação" autocomplete="off" required>
                        </div>
                        <div class="md:col-span-2">
                            <label class="block text-sm font-medium">Obra</label>
                            <select id="req-obra" class="w-full mt-1 p-2 border border-slate-200 rounded focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" required>
                                <option value="TABOCA">TABOCA</option>
                                <option value="ESTRELA" selected>ESTRELA</option>
                            </select>
                        </div>
                    </div>
                    <div class="md:col-span-2 flex items-start gap-3 p-3 rounded-lg bg-slate-50 border border-slate-200 mb-2">
                        <input type="checkbox" id="req-electronic-signature" class="mt-1 h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 shrink-0" checked>
                        <div class="min-w-0">
                            <label for="req-electronic-signature" class="text-sm font-semibold text-slate-800 cursor-pointer">Usar assinatura eletrônica (tablet)</label>
                            <p class="text-xs text-slate-500 mt-1 leading-relaxed">Marcado: assinar na tela antes de salvar. Desmarcado: o PDF sai com linhas em branco para assinar no papel.</p>
                        </div>
                    </div>
                    <div id="req-signature-section" class="mt-4 space-y-4">
                        <p class="text-sm text-slate-600">Desenhe abaixo com dedo ou caneta — obrigatório quando a assinatura eletrônica está ativada.</p>
                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
                            <div class="border border-slate-200 rounded-lg p-3 bg-white">
                                <p class="text-sm font-semibold text-slate-800 mb-2">Requisitante</p>
                                <canvas id="req-signature-requester-canvas" class="rounded border border-slate-300 bg-white w-full block"></canvas>
                                <button type="button" id="req-signature-requester-clear" class="mt-2 text-sm text-indigo-600 font-semibold hover:text-indigo-800">Limpar</button>
                            </div>
                            <div class="border border-slate-200 rounded-lg p-3 bg-white">
                                <p class="text-sm font-semibold text-slate-800 mb-2">Almoxarife</p>
                                <canvas id="req-signature-storekeeper-canvas" class="rounded border border-slate-300 bg-white w-full block"></canvas>
                                <button type="button" id="req-signature-storekeeper-clear" class="mt-2 text-sm text-indigo-600 font-semibold hover:text-indigo-800">Limpar</button>
                            </div>
                        </div>
                    </div>
                    <h3 class="font-semibold mt-6 mb-2">Itens</h3>
                    <div class="mb-2">
                        <input type="text" id="req-item-search" class="w-full p-2 border border-slate-200 rounded" placeholder="Pesquisar item (nome, código, local, grupo)...">
                    </div>
                    <div id="req-items-container" class="space-y-2 max-h-60 overflow-y-auto p-2 bg-slate-50 border rounded-md"></div>
                    <button type="button" id="add-req-item-btn" class="mt-2 text-sm text-indigo-600 font-semibold hover:text-indigo-800">+ Adicionar Item</button>
                    <div class="mt-8 flex justify-end gap-4 flex-wrap">
                        <button type="button" class="close-modal-btn px-6 py-2 bg-slate-200 rounded-lg font-semibold hover:bg-slate-300 transition">Cancelar</button>
                        <button type="submit" id="req-submit-stock" class="px-6 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition">Salvar e Baixar Estoque</button>
                        <button type="submit" id="req-submit-save-print" class="px-6 py-2 bg-slate-700 text-white rounded-lg font-semibold hover:bg-slate-800 transition">Salvar e imprimir</button>
                    </div>
                </form>
            `;
            document.getElementById('generic-modal').innerHTML = content;
            document.getElementById('generic-modal').classList.replace('max-w-md', 'max-w-4xl');
            openModal('generic-modal');

            const reqItemsContainer = document.getElementById('req-items-container');
            const reqItemSearch = document.getElementById('req-item-search');

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

            const syncReqSignatureSectionVisibility = () => {
                const on = document.getElementById('req-electronic-signature')?.checked ?? true;
                document.getElementById('req-signature-section')?.classList.toggle('hidden', !on);
            };
            document.getElementById('req-electronic-signature')?.addEventListener('change', syncReqSignatureSectionVisibility);
            syncReqSignatureSectionVisibility();

            requestAnimationFrame(() => {
                setupSignaturePad(document.getElementById('req-signature-requester-canvas'), document.getElementById('req-signature-requester-clear'));
                setupSignaturePad(document.getElementById('req-signature-storekeeper-canvas'), document.getElementById('req-signature-storekeeper-clear'));
            });
        };

        const buildRequisitionPrintPagesHtml = (reqs) => {
            const escHtml = (s) => String(s ?? '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;');

            let allContentHTML = '';
            reqs.forEach((req, index) => {
                let itemsHTML = (req.items || []).map(item => `
                    <tr>
                        <td class="border p-2">${item.productCodeRM || item.productCode}</td>
                        <td class="border p-2">${item.productName}</td>
                        <td class="border p-2 text-center">${item.quantity}</td>
                    </tr>
                `).join('');

                const logoHTML = appSettings.logoUrl ? `<img src="${appSettings.logoUrl}" class="h-12 w-auto max-w-[150px] object-contain">` : '';
                const pageBreak = index < reqs.length - 1 ? 'page-break-after: always;' : '';

                const useManualSignaturePdf = req.signatureMode === 'manual'
                    || (!req.requesterSignatureJpeg && !req.storekeeperSignatureJpeg);
                let signatureBlockHtml = '';
                if (useManualSignaturePdf) {
                    signatureBlockHtml = `
                         <div class="mt-12 text-sm">
                              <p class="text-xs text-slate-600 mb-3 text-center">Assinaturas manuais — preencher no papel após impressão</p>
                              <div class="flex flex-wrap justify-center gap-10 items-end">
                                  <div class="text-center" style="min-width:200px;max-width:48%;">
                                      <div style="height:96px;border:1px dashed #94a3b8;border-radius:8px;background:#fafafa;"></div>
                                      <p class="border-t border-black pt-1 px-4 mt-2">Assinatura requisitante</p>
                                      <p class="text-xs text-slate-600 mt-1">${escHtml(req.requester || '')}</p>
                                  </div>
                                  <div class="text-center" style="min-width:200px;max-width:48%;">
                                      <div style="height:96px;border:1px dashed #94a3b8;border-radius:8px;background:#fafafa;"></div>
                                      <p class="border-t border-black pt-1 px-4 mt-2">Assinatura almoxarife</p>
                                      <p class="text-xs text-slate-600 mt-1">${escHtml(req.storekeeperSignedBy || '________________')}</p>
                                  </div>
                              </div>
                         </div>`;
                } else {
                    signatureBlockHtml = `
                         <div class="mt-12 text-sm">
                              <div class="flex flex-wrap justify-center gap-10 items-end">
                                  <div class="text-center" style="min-width:200px;max-width:48%;">
                                      ${req.requesterSignatureJpeg ? `<img src="${req.requesterSignatureJpeg.replace(/"/g, '&quot;')}" alt="" style="max-height:100px;object-fit:contain;display:block;margin:0 auto 8px;">` : '<div style="height:80px;"></div>'}
                                      <p class="border-t border-black pt-1 px-4">Assinatura requisitante</p>
                                      <p class="text-xs text-slate-600 mt-1">${escHtml(req.requester || '')}</p>
                                  </div>
                                  <div class="text-center" style="min-width:200px;max-width:48%;">
                                      ${req.storekeeperSignatureJpeg ? `<img src="${req.storekeeperSignatureJpeg.replace(/"/g, '&quot;')}" alt="" style="max-height:100px;object-fit:contain;display:block;margin:0 auto 8px;">` : '<div style="height:80px;"></div>'}
                                      <p class="border-t border-black pt-1 px-4">Assinatura almoxarife</p>
                                      <p class="text-xs text-slate-600 mt-1">${escHtml(req.storekeeperSignedBy || '—')}</p>
                                  </div>
                              </div>
                         </div>`;
                }

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
                            <p><strong>Data:</strong> ${req.date?.seconds ? new Date(req.date.seconds * 1000).toLocaleDateString('pt-BR') : '...'}</p>
                            <p><strong>Requisitante:</strong> ${req.requester}</p>
                            <p><strong>Função do Funcionário:</strong> ${req.teamLeader || 'N/A'}</p>
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
                        ${signatureBlockHtml}
                    </div>
                `;
            });
            return allContentHTML;
        };

        const triggerRequisitionBrowserPrint = (reqs) => {
            if (!reqs || reqs.length === 0) return;
            const printArea = document.getElementById('print-area');
            if (!printArea) return;
            const html = buildRequisitionPrintPagesHtml(reqs);
            printArea.innerHTML = `<div id="pdf-content" class="bg-white p-4">${html}</div>`;
            printArea.classList.remove('hidden');
            const finalize = () => {
                printArea.classList.add('hidden');
                printArea.innerHTML = '';
                window.removeEventListener('afterprint', finalize);
            };
            window.addEventListener('afterprint', finalize);
            requestAnimationFrame(() => window.print());
        };

        const showViewRequisitionModal = (reqIds) => {
            const reqs = requisitions.filter(r => reqIds.includes(r.id));
            if (reqs.length === 0) return;

            const allContentHTML = buildRequisitionPrintPagesHtml(reqs);

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
                        <div class="md:col-span-2">
                            <label class="block text-sm font-medium text-slate-600">URL da foto (opcional)</label>
                            <input type="url" id="edit-product-photo-url" class="w-full mt-1 p-3 border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500/50 focus:border-indigo-500" placeholder="https://..." value="${escapeHtmlAttr(p.imageUrl || '')}">
                            <p class="text-xs text-slate-500 mt-1">HTTPS recomendado. Se não carregar, o site pode bloquear hotlink.</p>
                            <div id="edit-photo-preview-wrap" class="mt-2 hidden">
                                <img id="edit-photo-preview" alt="Prévia" class="max-h-28 rounded-lg border border-slate-200 object-contain bg-slate-50">
                            </div>
                        </div>
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
            updateProductPhotoPreview('edit-product-photo-url', 'edit-photo-preview-wrap', 'edit-photo-preview');
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

        /** Link aberto ao escanear o QR da plaquinha (mesma origem + ?p=productId) */
        const buildProductPlaqueUrl = (productId) => {
            if (!productId) return `${window.location.origin}${window.location.pathname || '/'}`;
            const u = new URL(window.location.href);
            u.search = '';
            u.hash = '';
            u.searchParams.set('p', productId);
            return u.toString();
        };

        let _plaqueDeepLinkConsumed = false;
        const tryOpenPlaqueDeepLink = () => {
            if (_plaqueDeepLinkConsumed) return;
            let pid;
            try { pid = new URLSearchParams(window.location.search).get('p'); } catch { return; }
            if (!pid || !isDataLoaded) return;
            const product = products.find(p => p.id === pid);
            if (!product) {
                if (products.length > 0) {
                    _plaqueDeepLinkConsumed = true;
                    showToast('Produto não encontrado. Verifique o link da plaquinha.', true);
                    try {
                        const u = new URL(window.location.href);
                        u.searchParams.delete('p');
                        const q = u.searchParams.toString();
                        history.replaceState(null, '', u.pathname + (q ? `?${q}` : '') + (u.hash || ''));
                    } catch (e) { /* noop */ }
                }
                return;
            }
            _plaqueDeepLinkConsumed = true;
            showPlaqueInfoModal(product);
            try {
                const u = new URL(window.location.href);
                u.searchParams.delete('p');
                const q = u.searchParams.toString();
                const newUrl = u.pathname + (q ? `?${q}` : '') + (u.hash || '');
                history.replaceState(null, '', newUrl);
            } catch (e) { console.warn('replaceState p:', e); }
        };

        const showPlaqueInfoModal = (product) => {
            if (!product) return;
            const appLink = buildProductPlaqueUrl(product.id);
            const productHistory = history
                .filter(h => h.productId === product.id)
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0))
                .slice(0, 50);

            let historyItemsHTML = '';
            if (productHistory.length > 0) {
                productHistory.forEach(h => {
                    let details = '';
                    let colorClass = '';
                    switch (h.type) {
                        case 'Entrada': case 'Ajuste Entrada':
                            colorClass = 'border-green-500'; details = `<strong>+${h.quantity}</strong> unidades. Estoque: <strong>${h.newTotal}</strong>.`; break;
                        case 'Saída': case 'Saída por Requisição':
                            colorClass = 'border-red-500'; details = `<strong>-${h.quantity}</strong> unidades. ${h.withdrawnBy ? `Retirado: <strong>${h.withdrawnBy}</strong>. ` : ''}Estoque: <strong>${h.newTotal}</strong>.`; break;
                        case 'Ajuste Saída':
                            colorClass = 'border-yellow-500'; details = `<strong>-${h.quantity}</strong> ajuste. Estoque: <strong>${h.newTotal}</strong>.`; break;
                        case 'Criação': case 'Importação':
                            colorClass = 'border-blue-500'; details = `Cadastro: <strong>${h.quantity}</strong> unidades.`; break;
                        case 'Edição':
                            colorClass = 'border-purple-500'; details = `Alteração de cadastro.`; break;
                        default:
                            colorClass = 'border-slate-400'; details = String(h.type || '—');
                    }
                    historyItemsHTML += `
                        <div class="p-3 bg-slate-50 rounded-lg border-l-4 ${colorClass}">
                            <p class="font-semibold text-slate-800">${h.type || '—'}</p>
                            <p class="text-sm text-slate-700">${details}</p>
                            <p class="text-xs text-slate-500 mt-1">${h.date ? new Date(h.date.seconds * 1000).toLocaleString('pt-BR') : '—'}</p>
                        </div>`;
                });
            } else {
                historyItemsHTML = '<p class="text-slate-500 text-center py-4">Nenhum histórico de movimentação para este produto.</p>';
            }

            const isLow = (product.quantity || 0) <= (product.minQuantity || 0);
            const content = `
                <div class="flex justify-between items-start gap-2">
                    <div class="min-w-0">
                        <p class="text-xs font-bold uppercase tracking-wide text-indigo-600">Ficha do produto</p>
                        <h2 class="text-xl sm:text-2xl font-bold text-slate-800 break-words">${product.name || '—'}</h2>
                        <p class="text-sm text-slate-600 mt-1">RM: <span class="font-mono font-semibold">${product.codeRM || '—'}</span> · SKU: <span class="font-mono">${product.code || '—'}</span></p>
                    </div>
                    <button type="button" class="close-modal-btn p-2 -mt-1 text-slate-400 hover:text-slate-700 text-2xl flex-shrink-0" aria-label="Fechar">&times;</button>
                </div>
                <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 mt-4">
                    <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                        <p class="text-xs text-slate-500">Estoque atual</p>
                        <p class="text-lg font-extrabold text-slate-900">${product.quantity ?? 0}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                        <p class="text-xs text-slate-500">Estoque mínimo</p>
                        <p class="text-lg font-extrabold text-slate-900">${product.minQuantity ?? 0}</p>
                    </div>
                    <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                        <p class="text-xs text-slate-500">Unidade</p>
                        <p class="text-sm font-bold text-slate-800 truncate">${product.unit || '—'}</p>
                    </div>
                    <div class="p-3 rounded-xl ${isLow ? 'bg-amber-50 border-amber-200' : 'bg-slate-50 border-slate-200'} border col-span-2 sm:col-span-1">
                        <p class="text-xs text-slate-500">Local</p>
                        <p class="text-sm font-bold text-slate-800 break-words">${product.location || '—'}</p>
                    </div>
                </div>
                <div class="mt-4 p-3 rounded-xl bg-indigo-50/80 border border-indigo-100">
                    <p class="text-xs font-bold text-indigo-800 mb-2">Acesso ao app (o QR da plaquinha leva a este endereço)</p>
                    <a id="plaque-app-link" href="#" class="text-sm text-indigo-600 font-semibold break-all underline">Abrir ficha do produto no app</a>
                    <p class="text-xs text-slate-500 mt-2">Toque no link no celular para abrir o app. Se estiver deslogado, faça login; em seguida abra o link de novo (ou use o leitor de QR da plaquinha após o login).</p>
                </div>
                <div class="mt-4 flex flex-wrap gap-2">
                    <button type="button" id="plaque-go-inventory" class="px-4 py-2 bg-indigo-600 text-white text-sm font-bold rounded-lg hover:bg-indigo-700">Ver na lista de estoque</button>
                </div>
                <h3 class="text-sm font-bold text-slate-800 mt-6 mb-2">Histórico de entradas e saídas</h3>
                <div class="max-h-72 overflow-y-auto space-y-2 pr-1">${historyItemsHTML}</div>
            `;
            const gm = document.getElementById('generic-modal');
            gm.innerHTML = content;
            gm.classList.replace('max-w-md', 'max-w-4xl');
            const appA = document.getElementById('plaque-app-link');
            if (appA) appA.href = appLink;
            openModal('generic-modal');
            document.getElementById('plaque-go-inventory')?.addEventListener('click', () => {
                closeModal('generic-modal');
                searchInput.value = product.name || '';
                currentPage = 1;
                switchView('inventory-view');
                renderProducts();
            });
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


        // ══════════════════════════════════════════════════════════════════
        // PLAQUINHAS — listagem, seleção e geração de PDF
        // ══════════════════════════════════════════════════════════════════
        const updatePlaquesCounter = () => {
            const checked = document.querySelectorAll('.plaque-check:checked').length;
            const counter = document.getElementById('plaques-counter');
            const genBtn  = document.getElementById('generate-plaques-btn');
            const preBtn  = document.getElementById('preview-plaques-btn');
            if (counter) counter.textContent = `${checked} selecionado(s)`;
            [genBtn, preBtn].forEach(btn => {
                if (!btn) return;
                btn.disabled = checked === 0;
                btn.style.opacity = checked === 0 ? '0.5' : '1';
            });
        };

        const renderPlaquesView = () => {
            const list       = document.getElementById('plaques-product-list');
            const noMsg      = document.getElementById('no-plaques-message');
            const totalCount = document.getElementById('plaques-total-count');
            if (!list) return;
            list.innerHTML = '';

            if (products.length === 0) {
                noMsg?.classList.remove('hidden');
                if (totalCount) totalCount.textContent = '0 produtos';
                updatePlaquesCounter();
                return;
            }
            noMsg?.classList.add('hidden');
            if (totalCount) totalCount.textContent = `${products.length} produto(s)`;

            const sorted = [...products].sort((a, b) =>
                (a.name || '').localeCompare(b.name || '', 'pt-BR'));

            sorted.forEach(p => {
                const isLow  = (p.quantity || 0) <= (p.minQuantity || 0);
                const qtyClr = isLow ? 'background:#fee2e2;color:#b91c1c;' : 'background:#dcfce7;color:#15803d;';
                const safeName = (p.name || '').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
                const safeRM   = (p.codeRM || '').replace(/"/g, '&quot;');
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-blue-50/30 transition-colors cursor-pointer';
                tr.innerHTML = `
                    <td class="p-3">
                        <input type="checkbox" class="plaque-check h-4 w-4 rounded border-slate-300 accent-blue-600"
                            data-id="${p.id}" data-name="${safeName}" data-rm="${safeRM}">
                    </td>
                    <td class="p-3 font-semibold text-slate-800 text-sm">${p.name || '—'}</td>
                    <td class="p-3 text-slate-500 text-sm font-mono">${p.codeRM || '—'}</td>
                    <td class="p-3 text-center">
                        <span class="px-2.5 py-0.5 rounded-full text-xs font-bold" style="${qtyClr}">${p.quantity ?? 0}</span>
                    </td>
                    <td class="p-3 text-center text-xs text-slate-400">${p.minQuantity ?? 0}</td>`;
                tr.addEventListener('click', e => {
                    if (e.target.tagName === 'INPUT') return;
                    const chk = tr.querySelector('.plaque-check');
                    chk.checked = !chk.checked;
                    updatePlaquesCounter();
                });
                list.appendChild(tr);
            });

            // select-all wiring
            const selectAll = document.getElementById('plaques-select-all');
            if (selectAll) {
                selectAll.checked = false;
                selectAll.onchange = () => {
                    document.querySelectorAll('.plaque-check').forEach(c => c.checked = selectAll.checked);
                    updatePlaquesCounter();
                };
            }
            list.addEventListener('change', updatePlaquesCounter);
            updatePlaquesCounter();
        };

        // Carrega o logo GEL como base64 uma única vez para uso no PDF
        let _gelLogoBase64 = null;
        const loadGelLogo = () => new Promise(resolve => {
            if (_gelLogoBase64) { resolve(_gelLogoBase64); return; }
            const img = new Image();
            img.crossOrigin = 'anonymous';
            img.onload = () => {
                try {
                    const canvas = document.createElement('canvas');
                    canvas.width  = img.naturalWidth;
                    canvas.height = img.naturalHeight;
                    canvas.getContext('2d').drawImage(img, 0, 0);
                    _gelLogoBase64 = canvas.toDataURL('image/png');
                    resolve(_gelLogoBase64);
                } catch { resolve(null); }
            };
            img.onerror = () => resolve(null);
            img.src = 'assets/img/logo-gel.png?v=1';
        });

        // qrcode (npm) — o antigo <script> do jsdelivr apontava para arquivo inexistente (404).
        // Carregamos via import dinâmico (main.js é ES module).
        let _qrcodeLib = undefined;
        const loadQRCodeLib = async () => {
            if (_qrcodeLib !== undefined) return _qrcodeLib;
            const urls = [
                'https://esm.sh/qrcode@1.5.3',
                'https://esm.run/qrcode@1.5.3',
                'https://cdn.skypack.dev/qrcode@1.5.3'
            ];
            for (const url of urls) {
                try {
                    const m = await import(url);
                    const lib = m.default || m;
                    if (lib && typeof lib.toDataURL === 'function') {
                        _qrcodeLib = lib;
                        return _qrcodeLib;
                    }
                } catch (e) {
                    console.warn('QR lib import falhou:', url, e);
                }
            }
            _qrcodeLib = null;
            return null;
        };

        const buildPlaquesPDF = async () => {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });

            const perPage = parseInt(document.querySelector('input[name="plaques-per-page"]:checked')?.value || '10');
            const rows    = perPage === 10 ? 5 : 4;
            const cols    = 2;

            // ── Dimensões A4 ──────────────────────────────────────────────
            const pageW   = 210, pageH = 297;
            const marginX = 10,  marginY = 10;
            const gapX    = 5,   gapY   = 4;
            // cardW ≈ 92.5mm  |  cardH ≈ 52.2mm (10/pág) ou 65.75mm (8/pág)
            const cardW = (pageW - marginX * 2 - gapX * (cols - 1)) / cols;
            const cardH = (pageH - marginY * 2 - gapY * (rows - 1)) / rows;

            const selected = [...document.querySelectorAll('.plaque-check:checked')];
            if (selected.length === 0) return null;

            // ── Pré-carrega logo e QR codes em paralelo ───────────────────
            const logoB64    = await loadGelLogo();
            const logoAspect = 414 / 259;
            const logoH      = 10;
            const logoW      = logoH * logoAspect;  // ≈ 16mm
            const logoPadR   = 2.5;
            const logoPadT   = 2;

            // QR = texto curto (lido direto no leitor, sem abrir app)
            const QR = await loadQRCodeLib();
            const qrCache = {};
            if (QR) {
                for (const chk of selected) {
                    const pid = chk.getAttribute('data-id') || '';
                    if (!pid) continue;
                    const p = (Array.isArray(products) ? products : []).find(pp => pp?.id === pid) || null;

                    const qty  = Number.isFinite(Number(p?.quantity)) ? Number(p.quantity) : 0;
                    const min  = Number.isFinite(Number(p?.minQuantity)) ? Number(p.minQuantity) : 0;
                    const qrText = `ATUAL: ${qty}\nMIN: ${min}`;

                    if (!qrCache[qrText]) {
                        try {
                            qrCache[qrText] = await QR.toDataURL(qrText, {
                                width: 400,
                                margin: 0,
                                errorCorrectionLevel: 'M',
                                color: { dark: '#0F172A', light: '#FFFFFF' }
                            });
                        } catch (e) {
                            console.warn('toDataURL QR falhou:', e);
                            qrCache[qrText] = null;
                        }
                    }
                }
            } else {
                console.warn('Biblioteca qrcode não disponível — plaquinhas sem QR.');
            }

            const footerH = 15;  // mm — rodapé (RM + QR maior)
            const qrSize  = 12.5; // mm — QR no PDF

            selected.forEach((chk, i) => {
                const posOnPage = i % perPage;
                if (i > 0 && posOnPage === 0) doc.addPage();

                const col = posOnPage % cols;
                const row = Math.floor(posOnPage / cols);
                const x   = marginX + col * (cardW + gapX);
                const y   = marginY + row * (cardH + gapY);

                const name   = chk.getAttribute('data-name') || '';
                const rm     = chk.getAttribute('data-rm') || '';
                const pid    = chk.getAttribute('data-id') || '';
                const p = (Array.isArray(products) ? products : []).find(pp => pp?.id === pid) || null;

                const resolvedRm   = (p?.codeRM || rm || p?.code || '').toString().trim();
                const qty          = Number.isFinite(Number(p?.quantity)) ? Number(p.quantity) : 0;
                const min          = Number.isFinite(Number(p?.minQuantity)) ? Number(p.minQuantity) : 0;
                const qrText = `ATUAL: ${qty}\nMIN: ${min}`;

                const qrB64  = qrText ? (qrCache[qrText] || null) : null;

                // ── Fundo branco + borda cinza ────────────────────────────
                doc.setFillColor(255, 255, 255);
                doc.setDrawColor(180, 180, 180);
                doc.setLineWidth(0.3);
                doc.roundedRect(x, y, cardW, cardH, 2, 2, 'FD');

                // ── Logo GEL — canto superior esquerdo ───────────────────
                if (logoB64) {
                    doc.addImage(logoB64, 'PNG', x + logoPadR, y + logoPadT, logoW, logoH);
                }

                // ── Nome do produto (auto-size) ───────────────────────────
                const len      = name.length;
                const fontSize = len > 45 ? 6.5 : len > 35 ? 7.5 : len > 24 ? 9 : len > 14 ? 11 : 12;
                doc.setFont('helvetica', 'bold');
                doc.setFontSize(fontSize);
                doc.setTextColor(15, 23, 42);

                const nameMaxW  = cardW - 8;
                const nameLines = doc.splitTextToSize(name, nameMaxW);
                const lineH     = fontSize * 0.38;
                const blockH    = nameLines.length * lineH;

                const nameAreaTop    = y + 3;
                const nameAreaBottom = (resolvedRm || qrText) ? y + cardH - footerH - 1 : y + cardH - 3;
                const nameAreaH      = nameAreaBottom - nameAreaTop;
                const nameStartY     = nameAreaTop + (nameAreaH - blockH) / 2 + lineH;

                doc.text(nameLines, x + cardW / 2, nameStartY, { align: 'center', lineHeightFactor: 1.3 });

                // ── Rodapé: separador + RM (esquerda) + QR (direita) ─────
                if (resolvedRm || qrText) {
                    const sepY = y + cardH - footerH;

                    // Linha separadora
                    doc.setDrawColor(210, 215, 225);
                    doc.setLineWidth(0.2);
                    doc.line(x + 3, sepY, x + cardW - 3, sepY);

                    doc.setFont('helvetica', 'normal');
                    doc.setFontSize(6.5);
                    doc.setTextColor(100, 116, 139);
                    // RM + linha “escaneie p/ ficha no app”
                    if (resolvedRm) {
                        doc.setFont('helvetica', 'bold');
                        doc.setFontSize(7);
                        if (qrB64) {
                            doc.text(`RM: ${resolvedRm}`, x + 4, y + cardH - footerH + 3.2, { align: 'left' });
                        } else {
                            doc.text(`RM: ${resolvedRm}`, x + cardW / 2, y + cardH - (footerH / 2) + 1, { align: 'center' });
                        }
                    }
                    if (qrB64) {
                        doc.setFont('helvetica', 'normal');
                        doc.setFontSize(5.5);
                        doc.setTextColor(140, 140, 140);
                        const note = 'Escaneie o QR p/ ver informações (sem abrir o app)';
                        const noteW = cardW - qrSize - 8;
                        const noteLines = doc.splitTextToSize(note, noteW);
                        doc.text(noteLines, x + 4, y + cardH - footerH + (resolvedRm ? 6.2 : 3.2), { lineHeightFactor: 1.15 });
                    }

                    // QR Code — canto inferior direito do rodapé
                    if (qrB64) {
                        const qrX = x + cardW - qrSize - 1;
                        const qrY = y + cardH - qrSize - 1;
                        try {
                            doc.addImage(qrB64, 'PNG', qrX, qrY, qrSize, qrSize);
                        } catch (e) {
                            const b64 = qrB64.replace(/^data:image\/\w+;base64,/, '');
                            try {
                                doc.addImage(b64, 'PNG', qrX, qrY, qrSize, qrSize);
                            } catch (e2) {
                                console.warn('addImage QR no PDF:', e, e2);
                            }
                        }
                    }
                }
            });

            return doc;
        };

        document.getElementById('generate-plaques-btn')?.addEventListener('click', async () => {
            const doc = await buildPlaquesPDF();
            if (doc) {
                doc.save('plaquinhas_estoque.pdf');
                showToast('PDF de plaquinhas gerado com sucesso!');
            }
        });

        document.getElementById('preview-plaques-btn')?.addEventListener('click', async () => {
            const doc = await buildPlaquesPDF();
            if (doc) window.open(doc.output('bloburl'), '_blank');
        });

        const renderComprasView = (periodDays = 15) => {
            const now = Date.now();
            const cutoff = now - periodDays * 864e5;

            // ── #3: cobertura alvo configurável ─────────────────────────────
            const coverageDays = parseInt(document.getElementById('compras-coverage-select')?.value || '30');

            // ── #7: consumo + último movimento por produto ───────────────────
            const consumoMap = {};
            const lastMovMap = {}; // #7: timestamp da última saída (todo histórico)
            history.forEach(h => {
                if (!h.productId) return;
                const isExit = h.type === 'Saída' || h.type === 'Saída por Requisição' || h.type === 'Ajuste Saída';
                if (!isExit) return;
                const ts = h.date?.seconds ? h.date.seconds * 1000 : 0;
                // atualiza último movimento independente do período
                if (!lastMovMap[h.productId] || ts > lastMovMap[h.productId]) lastMovMap[h.productId] = ts;
                // soma consumo apenas dentro do período
                if (ts >= cutoff) consumoMap[h.productId] = (consumoMap[h.productId] || 0) + Math.abs(h.quantity || 0);
            });

            // ── #5: popular select de grupos ────────────────────────────────
            const groupSelect = document.getElementById('compras-group-filter');
            if (groupSelect) {
                const preserved = groupSelect.value;
                const groups = [...new Set(products.map(p => p.group || 'Sem Grupo'))].sort();
                groupSelect.innerHTML = '<option value="">Todos os grupos</option>' +
                    groups.map(g => `<option value="${g}"${g === preserved ? ' selected' : ''}>${g}</option>`).join('');
            }

            const activeFilter  = document.querySelector('.compras-filter-btn[style*="background:#191c1d"]')?.dataset.filter || 'all';
            const searchQuery   = (document.getElementById('compras-search')?.value || '').trim().toLowerCase(); // #6
            const groupFilter   = document.getElementById('compras-group-filter')?.value || '';                   // #5

            // ── Montar itens ─────────────────────────────────────────────────
            const items = products.map(p => {
                const qty          = p.quantity || 0;
                const min          = p.minQuantity || 0;
                const totalConsumed = consumoMap[p.id] || 0;
                const dailyRate    = totalConsumed / periodDays;
                const daysRemaining = dailyRate > 0 ? Math.floor(qty / dailyRate) : null;
                const lastMov      = lastMovMap[p.id] || null; // #7

                // ── #3: qty sugerida com cobertura configurável ──────────────
                // ── #1: se sem histórico mas abaixo do mínimo → preenche até mín ──
                const coverageQty  = dailyRate > 0 ? Math.ceil(dailyRate * coverageDays * 1.2) : 0;
                const suggestedQty = dailyRate > 0
                    ? Math.max(0, coverageQty - qty)
                    : Math.max(0, min - qty); // #1: sem consumo → comprar diferença até mínimo

                let status;
                if (qty <= 0)                                         status = 'critico';
                else if (qty <= min)                                   status = 'critico';
                else if (daysRemaining !== null && daysRemaining <= 15) status = 'atencao';
                else if (qty <= min * 1.5)                             status = 'atencao';
                else if (totalConsumed > 0)                            status = 'monitorar';
                else return null; // estoque ok + sem consumo → não exibir

                return { p, totalConsumed, dailyRate, daysRemaining, suggestedQty, status, lastMov };
            }).filter(Boolean);

            // Ordenar: urgência → maior rotatividade (saída no período) → menos dias de cobertura
            const ordem = { critico: 0, atencao: 1, monitorar: 2 };
            items.sort((a, b) => {
                if (ordem[a.status] !== ordem[b.status]) return ordem[a.status] - ordem[b.status];
                if ((b.totalConsumed || 0) !== (a.totalConsumed || 0)) return (b.totalConsumed || 0) - (a.totalConsumed || 0);
                return (a.daysRemaining ?? 9999) - (b.daysRemaining ?? 9999);
            });

            const enrichComprasRow = (p) => {
                const qty = p.quantity || 0;
                const min = p.minQuantity || 0;
                const totalConsumed = consumoMap[p.id] || 0;
                const dailyRate = periodDays > 0 ? totalConsumed / periodDays : 0;
                const coverageQty = dailyRate > 0 ? Math.ceil(dailyRate * coverageDays * 1.2) : 0;
                const suggestedQty = dailyRate > 0
                    ? Math.max(0, coverageQty - qty)
                    : Math.max(0, min - qty);
                return { p, qty, min, totalConsumed, dailyRate, suggestedQty };
            };

            const turnoverRows = products
                .map((p) => enrichComprasRow(p))
                .filter((r) => r.totalConsumed > 0)
                .sort((a, b) => b.totalConsumed - a.totalConsumed)
                .slice(0, 40);

            const zeroRows = products
                .filter((p) => (p.quantity || 0) === 0)
                .map(enrichComprasRow)
                .sort((a, b) => b.totalConsumed - a.totalConsumed);

            const lowRows = products
                .filter((p) => {
                    const q = p.quantity || 0;
                    const m = p.minQuantity || 0;
                    return q > 0 && m > 0 && q <= m;
                })
                .map(enrichComprasRow)
                .sort((a, b) => {
                    const ra = a.min > 0 ? a.qty / a.min : 1;
                    const rb = b.min > 0 ? b.qty / b.min : 1;
                    return ra - rb;
                });

            lastComprasExport = {
                coverageDays,
                periodDays,
                pedido: items.filter((i) => i.suggestedQty > 0),
                turnover: turnoverRows,
                zero: zeroRows,
                low: lowRows
            };

            const comprasEscTxt = (v) => String(v ?? '')
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');

            const ttb = document.getElementById('compras-turnover-tbody');
            const ztb = document.getElementById('compras-zero-tbody');
            const ltb = document.getElementById('compras-low-tbody');
            if (ttb) {
                ttb.innerHTML = turnoverRows.length
                    ? turnoverRows.map((r, idx) => `
                    <tr class="hover:bg-slate-50">
                        <td class="px-2 py-2 font-bold text-slate-500">${idx + 1}</td>
                        <td class="px-2 py-2">
                            <span class="font-semibold text-slate-800">${comprasEscTxt(r.p.name)}</span>
                            <span class="block text-[10px] text-slate-400">${comprasEscTxt(r.p.codeRM || r.p.code || '—')}</span>
                        </td>
                        <td class="px-2 py-2 text-right font-bold text-indigo-700">${r.totalConsumed}</td>
                        <td class="px-2 py-2 text-right text-slate-600">${r.dailyRate > 0 ? r.dailyRate.toFixed(1) : '—'}</td>
                        <td class="px-2 py-2 text-right ${r.qty <= 0 ? 'text-red-600 font-bold' : 'text-slate-800'}">${r.qty}</td>
                    </tr>`).join('')
                    : '<tr><td colspan="5" class="px-3 py-6 text-center text-slate-400">Sem saídas registradas no período</td></tr>';
            }
            if (ztb) {
                ztb.innerHTML = zeroRows.length
                    ? zeroRows.map((r) => `
                    <tr class="hover:bg-slate-50">
                        <td class="px-2 py-2">
                            <span class="font-semibold text-slate-800">${comprasEscTxt(r.p.name)}</span>
                            <span class="block text-[10px] text-slate-400">${comprasEscTxt(r.p.codeRM || r.p.code || '—')}</span>
                        </td>
                        <td class="px-2 py-2 text-right text-slate-600">${r.min}</td>
                        <td class="px-2 py-2 text-right font-bold text-emerald-700">${r.suggestedQty > 0 ? r.suggestedQty : '—'}</td>
                    </tr>`).join('')
                    : '<tr><td colspan="3" class="px-3 py-6 text-center text-slate-400">Nenhum item com estoque zerado</td></tr>';
            }
            if (ltb) {
                ltb.innerHTML = lowRows.length
                    ? lowRows.map((r) => `
                    <tr class="hover:bg-slate-50">
                        <td class="px-2 py-2">
                            <span class="font-semibold text-slate-800">${comprasEscTxt(r.p.name)}</span>
                            <span class="block text-[10px] text-slate-400">${comprasEscTxt(r.p.codeRM || r.p.code || '—')}</span>
                        </td>
                        <td class="px-2 py-2 text-right font-semibold text-amber-800">${r.qty}</td>
                        <td class="px-2 py-2 text-right text-slate-600">${r.min}</td>
                        <td class="px-2 py-2 text-right font-bold text-emerald-700">${r.suggestedQty > 0 ? r.suggestedQty : '—'}</td>
                    </tr>`).join('')
                    : '<tr><td colspan="4" class="px-3 py-6 text-center text-slate-400">Nenhum item abaixo do mínimo (com mín. cadastrado)</td></tr>';
            }

            // Cards de resumo (baseados em TODOS os itens, antes dos filtros de UI)
            const criticos       = items.filter(i => i.status === 'critico').length;
            const atencoes       = items.filter(i => i.status === 'atencao').length;
            const monitorarCount = items.filter(i => i.status === 'monitorar').length;
            const totalSugerido  = items.reduce((acc, i) => acc + i.suggestedQty, 0);

            const summaryEl = document.getElementById('compras-summary');
            if (summaryEl) summaryEl.innerHTML = `
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
                    <p class="text-3xl font-black mt-1" style="color:#414754;">${monitorarCount}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">consumo ativo</p>
                </div>
                <div class="rounded-2xl p-4" style="background:#e8f5e9; border:1px solid #a5d6a7;">
                    <p class="text-xs font-bold uppercase tracking-wider" style="color:#006e2c;">📦 Para Comprar</p>
                    <p class="text-3xl font-black mt-1" style="color:#006e2c;">${criticos + atencoes}</p>
                    <p class="text-xs mt-0.5" style="color:#727785;">sugerido: ${totalSugerido} unid. / ${coverageDays}d</p>
                </div>
            `;

            // ── Aplicar filtros de UI (status + busca + grupo) ────────────────
            let filtered = activeFilter === 'all' ? items : items.filter(i => i.status === activeFilter);

            // #6: busca por nome ou código
            if (searchQuery) {
                filtered = filtered.filter(i =>
                    (i.p.name || '').toLowerCase().includes(searchQuery) ||
                    (i.p.codeRM || i.p.code || '').toLowerCase().includes(searchQuery)
                );
            }
            // #5: filtro por grupo
            if (groupFilter) {
                filtered = filtered.filter(i => (i.p.group || 'Sem Grupo') === groupFilter);
            }

            lastComprasFilteredItems = filtered;
            lastComprasPeriodDays    = periodDays;

            // Contador de resultados
            const countEl = document.getElementById('compras-count');
            if (countEl) countEl.textContent = filtered.length ? `${filtered.length} item(s) exibido(s)` : '';

            const tbody = document.getElementById('compras-list');
            const noMsg = document.getElementById('no-compras-message');
            tbody.innerHTML = '';

            if (filtered.length === 0) {
                noMsg.classList.remove('hidden');
                return;
            }
            noMsg.classList.add('hidden');

            filtered.forEach(({ p, dailyRate, daysRemaining, suggestedQty, status, lastMov, totalConsumed }) => {
                let badgeBg, badgeColor, badgeLabel;
                if (status === 'critico')      { badgeBg = '#fff2f0'; badgeColor = '#ba1a1a'; badgeLabel = '🔴 CRÍTICO'; }
                else if (status === 'atencao') { badgeBg = '#fff8e1'; badgeColor = '#795900'; badgeLabel = '🟡 ATENÇÃO'; }
                else                           { badgeBg = '#f3f4f5'; badgeColor = '#414754'; badgeLabel = '🔵 MONITORAR'; }

                const qty = p.quantity || 0;
                const min = p.minQuantity || 0;

                const daysText = daysRemaining !== null
                    ? (daysRemaining === 0
                        ? '<span style="color:#ba1a1a;font-weight:800;">ESGOTADO</span>'
                        : `<span style="color:${daysRemaining <= 7 ? '#ba1a1a' : daysRemaining <= 15 ? '#795900' : '#006e2c'};font-weight:700;">${daysRemaining}d</span>`)
                    : '<span style="color:#c1c6d6;">—</span>';

                const dailyText = dailyRate > 0
                    ? dailyRate.toFixed(1)
                    : '<span style="color:#c1c6d6;">—</span>';

                // #7: último movimento formatado
                const lastMovText = lastMov
                    ? (() => {
                        const d = new Date(lastMov);
                        const diffDays = Math.floor((now - lastMov) / 864e5);
                        const label = diffDays === 0 ? 'hoje' : diffDays === 1 ? 'ontem' : `${diffDays}d atrás`;
                        return `<span title="${d.toLocaleDateString('pt-BR')}" style="color:${diffDays <= 7 ? '#006e2c' : diffDays <= 30 ? '#795900' : '#727785'};font-size:11px;font-weight:600;">${label}</span>`;
                      })()
                    : '<span style="color:#c1c6d6;font-size:11px;">Sem registro</span>';

                // #5: grupo
                const grupo = p.group || '—';

                const tr = document.createElement('tr');
                tr.className = 'transition-colors hover:bg-slate-50';
                tr.innerHTML = `
                    <td class="px-4 py-3">
                        <span class="px-2 py-1 rounded-full text-xs font-bold whitespace-nowrap" style="background:${badgeBg}; color:${badgeColor};">${badgeLabel}</span>
                    </td>
                    <td class="px-4 py-3">
                        <p class="font-bold text-slate-800" style="font-size:13px;">${p.name}</p>
                        <p class="text-xs text-slate-400">${p.codeRM || p.code || '—'} · ${p.location || 'N/A'}</p>
                    </td>
                    <td class="px-4 py-3">
                        <span class="text-xs font-medium px-2 py-0.5 rounded-full" style="background:#f3f4f5; color:#414754;">${grupo}</span>
                    </td>
                    <td class="px-4 py-3 text-center font-bold" style="color:${qty <= min ? '#ba1a1a' : '#191c1d'};">
                        ${qty} <span class="text-xs font-normal text-slate-400">${p.unit || ''}</span>
                    </td>
                    <td class="px-4 py-3 text-center text-slate-500">${min}</td>
                    <td class="px-4 py-3 text-center font-semibold text-indigo-800">${totalConsumed || 0}</td>
                    <td class="px-4 py-3 text-center text-slate-600">${dailyText}</td>
                    <td class="px-4 py-3 text-center">${daysText}</td>
                    <td class="px-4 py-3 text-center">${lastMovText}</td>
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
            if (viewId === 'activity-log-view') renderActivityLog(activityLogSearchInput?.value || '');
            if (viewId === 'rm-view') renderRMView(rmSearchInput.value); 
            if (viewId === 'dashboard-view') updateDashboard();
            if (viewId === 'tool-loans-view') {
                renderToolLoanQueue();
                renderToolLoans();
            }
            if (viewId === 'plaques-view') renderPlaquesView();
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
            // Pega o valor do input hidden (ou select se ainda existir)
            const selectedObraId = loginObraSelect ? loginObraSelect.value : 'uhe_estrela';
            loginError.classList.add('hidden');

            console.log('Tentando login:', { username, selectedObraId });

            if (!username || !password) {
                loginError.textContent = 'Preencha o nome de usuário e a senha.';
                loginError.classList.remove('hidden');
                return;
            }

            loginBtn.disabled = true;
            loginBtn.textContent = 'Verificando...';

            try {
                const hash = await hashPassword(password);
                const userId = normalizeUserId(username);

                console.log('UserId normalizado:', userId);

                // 1. Verificar credenciais hardcoded
                const BUILTIN_USERS = {
                    'uhe_estrela': { displayName: 'UHE ESTRELA', role: 'admin', pwd: '60218', obraId: 'uhe_estrela' },
                    'pch_taboca': { displayName: 'PCH TABOCA', role: 'admin', pwd: '60218', obraId: 'pch_taboca' }
                };
                const builtin = BUILTIN_USERS[userId];

                if (builtin) {
                    const expectedHash = await hashPassword(builtin.pwd);
                    if (hash === expectedHash) {
                        const appUser = { uid: userId, displayName: builtin.displayName, role: builtin.role, obraId: builtin.obraId };
                        localStorage.setItem('appUser', JSON.stringify(appUser));
                        console.log('Iniciando sessão para:', appUser);
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
                console.log('Buscando usuário no Firestore...');
                const userRef = doc(db, `/artifacts/${appId}/public/data/users`, userId);
                const userDoc = await getDoc(userRef);

                if (userDoc.exists && userDoc.data().passwordHash === hash) {
                    const data = userDoc.data();
                    const appUser = { 
                        uid: userId, 
                        displayName: data.displayName || data.username || username, 
                        role: data.role || 'operador', 
                        obraId: selectedObraId 
                    };
                    localStorage.setItem('appUser', JSON.stringify(appUser));
                    initializeAppSession(appUser);
                } else {
                    loginError.textContent = userDoc.exists ? 'Senha incorreta.' : 'Usuário não encontrado.';
                    loginError.classList.remove('hidden');
                    loginBtn.disabled = false;
                    loginBtn.textContent = 'Entrar';
                }
            } catch (e) {
                console.error('Erro no login:', e);
                loginError.textContent = 'Erro ao fazer login: ' + e.message;
                loginError.classList.remove('hidden');
                loginBtn.disabled = false;
                loginBtn.textContent = 'Entrar';
            }
        };

        loginBtn.addEventListener('click', doLogin);
        loginPasswordInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') doLogin(); });
        loginUsernameInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') loginPasswordInput.focus(); });

        const setFeedbackPanelOpen = (open) => {
            if (!feedbackPanel) return;
            feedbackPanel.classList.toggle('hidden', !open);
            feedbackBackdrop?.classList.toggle('hidden', !open);
            if (feedbackBackdrop) feedbackBackdrop.setAttribute('aria-hidden', open ? 'false' : 'true');
            feedbackFooterBtn?.setAttribute('aria-expanded', open ? 'true' : 'false');
            if (open) document.body.style.overflow = 'hidden';
            else document.body.style.overflow = '';
        };

        logoutBtn.addEventListener('click', () => {
            stopEstrelaListeners();
            stopCoreListeners();
            setFeedbackPanelOpen(false);
            closeProductImageLightbox();
            feedbackForm?.reset();
            localStorage.removeItem('appUser');
            currentUser = null;
            userRole = null;
            isDataLoaded = false;
            isAuthInitialized = false;
            products = [];
            history = [];
            requisitions = [];
            toolLoans = [];
            toolLoanQueue = [];
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
        if (locationFilterSelect) {
            locationFilterSelect.addEventListener('change', (e) => {
                inventoryLocationFilter = e.target.value || 'all';
                currentPage = 1;
                renderProducts();
            });
        }

        toolLoanForm?.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!hasPermission('create')) {
                showToast('🔒 Você não tem permissão para cautelar ferramentas', true);
                return;
            }

            const borrower = toUpperText(document.getElementById('tool-loan-borrower').value);
            const role = toUpperText(document.getElementById('tool-loan-role').value);
            const observation = toUpperText(document.getElementById('tool-loan-observation').value);
            if (!borrower) {
                showToast('Informe o nome de quem pegou.', true);
                return;
            }

            let linesToProcess = [];
            if (toolLoanQueue.length > 0) {
                linesToProcess = toolLoanQueue.map((l) => ({ ...l }));
            } else {
                const productId = toolLoanProductSelect?.value;
                const quantity = parseInt(document.getElementById('tool-loan-quantity').value, 10);
                const product = products.find(p => p.id === productId);
                if (!productId || !product || isNaN(quantity) || quantity <= 0) {
                    showToast('Adicione itens à lista ou selecione uma ferramenta e quantidade.', true);
                    return;
                }
                if ((product.quantity || 0) < quantity) {
                    showToast(`Estoque insuficiente para ${product.name}. Disponível: ${product.quantity || 0}`, true);
                    return;
                }
                linesToProcess = [{
                    productId,
                    productName: product.name,
                    productCode: product.code || '',
                    productCodeRM: product.codeRM || '',
                    quantity
                }];
            }

            const qtyByProduct = {};
            linesToProcess.forEach((l) => {
                qtyByProduct[l.productId] = (qtyByProduct[l.productId] || 0) + l.quantity;
            });
            for (const pid of Object.keys(qtyByProduct)) {
                const p = products.find((x) => x.id === pid);
                if (!p || (p.quantity || 0) < qtyByProduct[pid]) {
                    showToast(`Estoque insuficiente ou item indisponível: ${p?.name || pid}`, true);
                    return;
                }
            }

            const submitBtn = document.getElementById('tool-loan-submit-btn') || e.target.querySelector('button[type="submit"]');
            submitBtn.disabled = true;
            if (toolLoanSubmitLabel) toolLoanSubmitLabel.textContent = 'Salvando...';

            try {
                const projectedQty = {};
                linesToProcess.forEach((l) => {
                    if (projectedQty[l.productId] === undefined) {
                        const pr = products.find((p) => p.id === l.productId);
                        projectedQty[l.productId] = pr ? (pr.quantity || 0) : 0;
                    }
                    projectedQty[l.productId] -= l.quantity;
                });
                for (const pid of Object.keys(projectedQty)) {
                    if (projectedQty[pid] < 0) {
                        showToast('Estoque insuficiente para concluir todas as linhas. Atualize a lista e tente de novo.', true);
                        return;
                    }
                }

                const batch = writeBatch(db);
                const createdBy = toUpperText(currentUser?.displayName || 'Anônimo');
                const productIdsUpdated = new Set();
                for (const line of linesToProcess) {
                    if (!productIdsUpdated.has(line.productId)) {
                        batch.update(doc(productsCollectionRef, line.productId), {
                            quantity: projectedQty[line.productId],
                            updatedAt: serverTimestamp()
                        });
                        productIdsUpdated.add(line.productId);
                    }
                    batch.set(doc(toolLoansCollectionRef), {
                        productId: line.productId,
                        productName: line.productName,
                        productCode: line.productCode || '',
                        productCodeRM: line.productCodeRM || '',
                        quantity: line.quantity,
                        borrower,
                        role: role || null,
                        observation: observation || null,
                        status: 'open',
                        loanDate: serverTimestamp(),
                        createdBy
                    });
                }
                await batch.commit();
                toolLoanForm.reset();
                document.getElementById('tool-loan-quantity').value = '1';
                toolLoanQueue = [];
                renderToolLoanQueue();
                const n = linesToProcess.length;
                showToast(n > 1 ? `${n} cautelas registradas e estoque atualizado.` : 'Ferramenta cautelada e estoque atualizado.');
            } catch (error) {
                console.error('Erro ao cautelar ferramenta:', error);
                showToast(`Erro ao cautelar: ${error.message}`, true);
            } finally {
                submitBtn.disabled = false;
                renderToolLoanQueue();
            }
        });

        toolLoanQueueBody?.addEventListener('click', (ev) => {
            const btn = ev.target.closest('.tool-loan-queue-remove');
            if (!btn) return;
            const idx = parseInt(btn.getAttribute('data-queue-idx'), 10);
            if (Number.isNaN(idx)) return;
            toolLoanQueue.splice(idx, 1);
            renderToolLoanQueue();
        });

        toolLoanAddQueueBtn?.addEventListener('click', () => {
            if (!hasPermission('create')) {
                showToast('🔒 Você não tem permissão para cautelar ferramentas', true);
                return;
            }
            const productId = toolLoanProductSelect?.value;
            const quantity = parseInt(document.getElementById('tool-loan-quantity')?.value, 10);
            const product = products.find((p) => p.id === productId);
            if (!productId || !product || !quantity || quantity < 1) {
                showToast('Selecione a ferramenta e uma quantidade válida.', true);
                return;
            }
            const existingInQueue = getToolLoanQueueQtyForProduct(productId);
            const avail = (product.quantity || 0) - existingInQueue;
            if (quantity > avail) {
                showToast(`Estoque insuficiente para ${product.name}. Disponível: ${avail} (considerando a lista).`, true);
                return;
            }
            const existingIdx = toolLoanQueue.findIndex((l) => l.productId === productId);
            if (existingIdx >= 0) {
                toolLoanQueue[existingIdx].quantity += quantity;
            } else {
                toolLoanQueue.push({
                    productId,
                    productName: product.name,
                    productCode: product.code || '',
                    productCodeRM: product.codeRM || '',
                    quantity
                });
            }
            renderToolLoanQueue();
            document.getElementById('tool-loan-quantity').value = '1';
            showToast('Item adicionado à lista.');
        });

        toolLoansListFilterBorrower?.addEventListener('input', () => renderToolLoans());
        toolLoansListFilterProduct?.addEventListener('input', () => renderToolLoans());
        toolLoansFilterReturnedStatus?.addEventListener('change', () => renderToolLoans());

        feedbackFooterBtn?.addEventListener('click', (ev) => {
            ev.stopPropagation();
            const willOpen = feedbackPanel?.classList.contains('hidden');
            setFeedbackPanelOpen(!!willOpen);
        });
        feedbackBackdrop?.addEventListener('click', () => setFeedbackPanelOpen(false));
        document.getElementById('feedback-close')?.addEventListener('click', () => setFeedbackPanelOpen(false));
        document.addEventListener('keydown', (ev) => {
            if (ev.key === 'Escape') setFeedbackPanelOpen(false);
        });
        feedbackForm?.addEventListener('submit', async (ev) => {
            ev.preventDefault();
            const msg = (feedbackMessage?.value || '').trim();
            if (msg.length < 5) {
                showToast('Escreva pelo menos 5 caracteres.', true);
                return;
            }
            const submitBtn = document.getElementById('feedback-submit');
            if (!submitBtn) return;
            submitBtn.disabled = true;
            try {
                await addDoc(collection(db, `/artifacts/${appId}/public/data/suggestions`), {
                    message: msg,
                    createdAt: serverTimestamp(),
                    viewId: currentViewId || null,
                    viewTitle: (currentViewTitle?.textContent || '').trim() || null,
                    userDisplay: currentUser?.displayName || null,
                    userUid: currentUser?.uid || null,
                    obraId: currentObraId || null,
                    userAgent: (navigator.userAgent || '').slice(0, 400),
                    appVersion: '2.0.3'
                });
                feedbackForm.reset();
                setFeedbackPanelOpen(false);
                showToast('Obrigado! Sua sugestão foi registrada.');
            } catch (err) {
                console.error(err);
                if (err.code === 'permission-denied') {
                    showToast('Permissão negada no Firebase. Peça ao admin para liberar gravação em suggestions.', true);
                } else {
                    showToast('Não foi possível enviar agora. Tente de novo.', true);
                }
            } finally {
                submitBtn.disabled = false;
            }
        });

        productList.addEventListener('click', async (e) => {
            const thumbBtn = e.target.closest('.product-thumb-lightbox');
            if (thumbBtn) {
                const url = thumbBtn.getAttribute('data-image-url');
                openProductImageLightbox(url);
                return;
            }

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
                    showStockDeleteConfirmationModal(
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

            showStockDeleteConfirmationModal(
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

            const photoRaw = document.getElementById('product-photo-url')?.value?.trim() || '';
            const photoSan = sanitizeProductImageUrl(photoRaw);
            if (photoRaw && !photoSan) {
                showToast('URL da foto inválida. Use http ou https com link direto da imagem.', true);
                return;
            }

            const newProduct = {
                code: document.getElementById('product-code').value.trim(),
                codeRM: document.getElementById('product-code-rm').value.trim(),
                name: toUpperText(document.getElementById('product-name').value),
                unit: document.getElementById('product-unit').value,
                group: document.getElementById('product-group').value,
                quantity: parseInt(document.getElementById('product-quantity').value),
                minQuantity: parseInt(document.getElementById('product-min-quantity').value),
                location: toUpperText(document.getElementById('product-location').value),
                observation: '',
                createdAt: serverTimestamp(),
                updatedAt: serverTimestamp(),
                createdBy: toUpperText(currentUser?.displayName || currentUser?.uid || 'Anônimo')
            };
            if (photoSan) newProduct.imageUrl = photoSan;

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
                document.getElementById('product-photo-preview-wrap')?.classList.add('hidden');
                document.getElementById('product-photo-preview')?.removeAttribute('src');
                switchView('inventory-view');
            } catch (error) {
                console.error("Erro ao adicionar produto: ", error);
                const errorMsg = error.code === 'permission-denied' 
                    ? '🔒 Sem permissão para adicionar produtos. Verifique as configurações do Firebase.'
                    : '❌ Falha ao adicionar produto. Tente novamente.';
                showToast(errorMsg, true);
            }
        });

        document.getElementById('product-photo-url')?.addEventListener('input', () => {
            updateProductPhotoPreview('product-photo-url', 'product-photo-preview-wrap', 'product-photo-preview');
        });
        document.getElementById('product-photo-url')?.addEventListener('blur', () => {
            updateProductPhotoPreview('product-photo-url', 'product-photo-preview-wrap', 'product-photo-preview');
        });
        document.body.addEventListener('input', (ev) => {
            if (ev.target?.id === 'edit-product-photo-url') {
                updateProductPhotoPreview('edit-product-photo-url', 'edit-photo-preview-wrap', 'edit-photo-preview');
            }
        });
        document.body.addEventListener('focusout', (ev) => {
            if (ev.target?.id === 'edit-product-photo-url') {
                updateProductPhotoPreview('edit-product-photo-url', 'edit-photo-preview-wrap', 'edit-photo-preview');
            }
        });

        document.getElementById('product-image-lightbox-close')?.addEventListener('click', () => closeProductImageLightbox());
        document.getElementById('product-image-lightbox-backdrop')?.addEventListener('click', () => closeProductImageLightbox());
        document.addEventListener('keydown', (ev) => {
            if (ev.key !== 'Escape') return;
            const lb = document.getElementById('product-image-lightbox');
            if (lb && !lb.classList.contains('hidden')) closeProductImageLightbox();
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
                        if (!productDoc.exists) {
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
                const formEl = e.target;
                const submitBtns = Array.from(formEl.querySelectorAll('button[type="submit"]'));
                const clickedSubmit = e.submitter || submitBtns[0];
                const wantsPrint = clickedSubmit?.id === 'req-submit-save-print';
                const resetReqSubmitBtns = () => {
                    submitBtns.forEach((b) => {
                        b.disabled = false;
                        if (b.id === 'req-submit-stock') b.textContent = 'Salvar e Baixar Estoque';
                        else if (b.id === 'req-submit-save-print') b.textContent = 'Salvar e imprimir';
                    });
                };
                submitBtns.forEach((b) => { b.disabled = true; });
                if (clickedSubmit) clickedSubmit.innerHTML = `<div class="spinner-small"></div>`;

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
                    resetReqSubmitBtns();
                    return;
                }

                const useElectronicSignature = document.getElementById('req-electronic-signature')?.checked ?? true;
                const newReqData = {
                    number: document.getElementById('req-number').value,
                    requester: toUpperText(document.getElementById('req-requester').value),
                    teamLeader: toUpperText(document.getElementById('req-team-leader').value),
                    applicationLocation: toUpperText(document.getElementById('req-application-location').value),
                    obra: document.getElementById('req-obra').value,
                    items: [],
                    date: serverTimestamp(),
                    signatureMode: useElectronicSignature ? 'electronic' : 'manual'
                };

                if (useElectronicSignature) {
                    const sigReqCanvas = document.getElementById('req-signature-requester-canvas');
                    const sigStoCanvas = document.getElementById('req-signature-storekeeper-canvas');
                    if (!sigReqCanvas || sigReqCanvas.dataset.hasInk !== '1') {
                        showToast('Assinatura do requisitante é obrigatória (modo eletrônico).', true);
                        resetReqSubmitBtns();
                        return;
                    }
                    if (!sigStoCanvas || sigStoCanvas.dataset.hasInk !== '1') {
                        showToast('Assinatura do almoxarife é obrigatória (modo eletrônico).', true);
                        resetReqSubmitBtns();
                        return;
                    }
                    newReqData.requesterSignatureJpeg = sigReqCanvas.toDataURL('image/jpeg', 0.85);
                    newReqData.storekeeperSignatureJpeg = sigStoCanvas.toDataURL('image/jpeg', 0.85);
                    newReqData.storekeeperSignedBy = toUpperText(currentUser?.displayName || currentUser?.uid || 'Anônimo');
                }

                try {
                    const batch = writeBatch(db);

                    itemsMap.forEach(({ quantity: requestedQuantity }, productId) => {
                        const productData = products.find(p => p.id === productId);
                        if (!productData) throw new Error(`Produto não encontrado na tela. Atualize o app e tente novamente.`);
                        if ((productData.quantity || 0) < requestedQuantity) throw new Error(`Estoque insuficiente para ${productData.name}.`);

                        const newQuantity = (productData.quantity || 0) - requestedQuantity;
                        batch.update(doc(productsCollectionRef, productId), { quantity: newQuantity, updatedAt: serverTimestamp() });

                        newReqData.items.push({
                            productId,
                            productName: productData.name,
                            productCode: productData.code,
                            productCodeRM: productData.codeRM,
                            quantity: requestedQuantity
                        });

                        batch.set(doc(historyCollectionRef), {
                            productId,
                            productCode: productData.code || 'N/A',
                            productCodeRM: productData.codeRM || 'N/A',
                            productName: toUpperText(productData.name || 'Produto Excluído'),
                            type: 'Saída por Requisição',
                            quantity: requestedQuantity,
                            newTotal: newQuantity,
                            withdrawnBy: newReqData.requester,
                            teamLeader: newReqData.teamLeader,
                            applicationLocation: newReqData.applicationLocation,
                            obra: newReqData.obra,
                            details: `Requisição Nº ${newReqData.number}`,
                            rmProcessed: false,
                            date: serverTimestamp(),
                            performedBy: toUpperText(currentUser?.displayName || 'Anônimo'),
                            timestamp: serverTimestamp()
                        });
                    });

                    const reqForPrint = {
                        number: newReqData.number,
                        requester: newReqData.requester,
                        teamLeader: newReqData.teamLeader,
                        applicationLocation: newReqData.applicationLocation,
                        obra: newReqData.obra,
                        items: [...newReqData.items],
                        signatureMode: newReqData.signatureMode,
                        date: { seconds: Math.floor(Date.now() / 1000) },
                        requesterSignatureJpeg: newReqData.requesterSignatureJpeg,
                        storekeeperSignatureJpeg: newReqData.storekeeperSignatureJpeg,
                        storekeeperSignedBy: newReqData.storekeeperSignedBy,
                    };

                    batch.set(doc(requisitionsCollectionRef), newReqData);
                    await batch.commit();
                    
                    selectedProductIds.clear();
                    showToast(wantsPrint
                        ? "Requisição salva. Confirme na janela de impressão para enviar à impressora."
                        : "Requisição criada e estoque atualizado!");
                    closeModal('generic-modal');
                    if (wantsPrint) {
                        triggerRequisitionBrowserPrint([reqForPrint]);
                    }

                } catch (error) {
                    console.error("Erro ao finalizar requisição: ", error);
                    const message = error.code === 'resource-exhausted'
                        ? 'Limite temporário do Firebase atingido. Aguarde alguns minutos e tente novamente.'
                        : error.message;
                    showToast(`Erro: ${message}`, true);
                } finally {
                    resetReqSubmitBtns();
                }
            }

            if (e.target.id === 'edit-product-form') {
                e.preventDefault();
                const id = document.getElementById('edit-product-id').value;
                const productRef = doc(productsCollectionRef, id);
                const newQuantity = parseInt(document.getElementById('edit-product-quantity').value);
                const photoRaw = document.getElementById('edit-product-photo-url')?.value?.trim() || '';
                const photoSan = sanitizeProductImageUrl(photoRaw);
                if (photoRaw && !photoSan) {
                    showToast('URL da foto inválida. Use http ou https com link direto da imagem.', true);
                    return;
                }
                const updatedData = {
                    codeRM: document.getElementById('edit-product-code-rm').value.trim(),
                    name: toUpperText(document.getElementById('edit-product-name').value),
                    unit: document.getElementById('edit-product-unit').value,
                    group: document.getElementById('edit-product-group').value,
                    quantity: newQuantity,
                    minQuantity: parseInt(document.getElementById('edit-product-min-quantity').value),
                    location: toUpperText(document.getElementById('edit-product-location').value),
                    observation: document.getElementById('edit-product-observation').value.trim(),
                    updatedAt: serverTimestamp(), // 🔍 Auditoria: timestamp de atualização
                    updatedBy: toUpperText(currentUser?.displayName || currentUser?.uid || 'Anônimo') // 🔍 Auditoria: quem atualizou
                };
                if (photoSan) updatedData.imageUrl = photoSan;
                else updatedData.imageUrl = deleteField();
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
                const name = toUpperText(document.getElementById('location-name').value);
                const code = document.getElementById('location-code').value.trim();
                const aliases = parseUpperAliases(document.getElementById('location-aliases').value);

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
                const aliases = parseUpperAliases(document.getElementById('edit-location-aliases').value);
                const updatedData = {
                    name: toUpperText(document.getElementById('edit-location-name').value),
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
                const supplier = toUpperText(document.getElementById('entry-supplier').value);
                const observation = toUpperText(document.getElementById('entry-observation').value);
                const receivedBy = toUpperText(currentUser?.displayName || currentUser?.uid || 'Sistema');

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
                        if (!productDoc.exists) {
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
            const returnToolLoanBtn = e.target.closest('.return-tool-loan-btn');
            if (returnToolLoanBtn) {
                const loanId = returnToolLoanBtn.dataset.id;
                const loan = toolLoans.find(item => item.id === loanId);
                if (!loan || loan.status === 'returned' || loan.status === 'damaged') return;
                if (!hasPermission('update')) {
                    showToast('🔒 Você não tem permissão para devolver cautelas', true);
                    return;
                }

                showConfirmationModal(
                    'Devolver Ferramenta',
                    `Confirmar devolução de ${loan.quantity} unidade(s) de ${loan.productName}?`,
                    async () => {
                        const product = products.find(p => p.id === loan.productId);
                        if (!product) {
                            showToast('Produto não encontrado no estoque. Atualize o app e tente novamente.', true);
                            return;
                        }

                        try {
                            const batch = writeBatch(db);
                            const newQuantity = (product.quantity || 0) + (loan.quantity || 0);
                            batch.update(doc(productsCollectionRef, loan.productId), { quantity: newQuantity, updatedAt: serverTimestamp() });
                            batch.update(doc(toolLoansCollectionRef, loan.id), {
                                status: 'returned',
                                returnDate: serverTimestamp(),
                                returnedBy: toUpperText(currentUser?.displayName || 'Anônimo')
                            });
                            await batch.commit();
                            showToast('Ferramenta devolvida ao estoque.');
                        } catch (error) {
                            console.error('Erro ao devolver ferramenta:', error);
                            showToast(`Erro ao devolver: ${error.message}`, true);
                        }
                    }
                );
            }

            const damagedToolLoanBtn = e.target.closest('.damaged-tool-loan-btn');
            if (damagedToolLoanBtn) {
                const loanId = damagedToolLoanBtn.dataset.id;
                const loan = toolLoans.find(item => item.id === loanId);
                if (!loan || loan.status === 'returned' || loan.status === 'damaged') return;
                if (!hasPermission('update')) {
                    showToast('🔒 Você não tem permissão para registrar dano', true);
                    return;
                }

                showConfirmationModal(
                    'Registrar Ferramenta Danificada',
                    `A ferramenta "${loan.productName}" será marcada como danificada e NÃO voltará ao estoque. Confirmar?`,
                    async () => {
                        try {
                            await updateDoc(doc(toolLoansCollectionRef, loan.id), {
                                status: 'damaged',
                                returnDate: serverTimestamp(),
                                returnedBy: toUpperText(currentUser?.displayName || 'Anônimo')
                            });
                            showToast('Ferramenta registrada como danificada. Estoque não alterado.');
                        } catch (error) {
                            console.error('Erro ao registrar dano:', error);
                            showToast(`Erro ao registrar: ${error.message}`, true);
                        }
                    }
                );
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

        exportToolLoansBtn?.addEventListener('click', () => {
            const XLS = typeof XLSX !== 'undefined' ? XLSX : null;
            if (!XLS) { showToast('Biblioteca de planilhas não carregou. Recarregue a página.', true); return; }
            if (toolLoans.length === 0) { showToast('Nenhuma cautela registrada para exportar.', true); return; }

            const today     = new Date().toLocaleDateString('pt-BR');
            const hora      = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
            const stamp     = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
            const wb        = XLS.utils.book_new();

            // ── ESTILOS ───────────────────────────────────────────────────────────
            const bdr = (color = 'CBD5E1') => ({
                top:    { style: 'thin', color: { rgb: color } },
                bottom: { style: 'thin', color: { rgb: color } },
                left:   { style: 'thin', color: { rgb: color } },
                right:  { style: 'thin', color: { rgb: color } }
            });
            const sTitle = (accent) => ({
                font:      { bold: true, sz: 18, color: { rgb: 'FFFFFF' } },
                fill:      { fgColor: { rgb: '0F172A' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border:    { bottom: { style: 'thick', color: { rgb: accent } } }
            });
            const sSub = {
                font:      { italic: true, sz: 10, color: { rgb: '94A3B8' } },
                fill:      { fgColor: { rgb: '1E293B' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
            const sSecHdr = (rgb) => ({
                font:      { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
                fill:      { fgColor: { rgb } },
                alignment: { horizontal: 'left', vertical: 'center' },
                border:    bdr(rgb)
            });
            const sTH = (rgb) => ({
                font:      { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
                fill:      { fgColor: { rgb } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border:    bdr(rgb)
            });
            const sCell = (even, align = 'left', color = '1E293B', bold = false) => ({
                font:      { sz: 10, color: { rgb: color }, bold },
                fill:      { fgColor: { rgb: even ? 'F8FAFC' : 'FFFFFF' } },
                alignment: { horizontal: align, vertical: 'center' },
                border:    bdr()
            });
            const sBadgeOpen = {
                font:  { bold: true, sz: 10, color: { rgb: 'D97706' } },
                fill:  { fgColor: { rgb: 'FEF9C3' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: bdr('D97706')
            };
            const sBadgeDone = {
                font:  { bold: true, sz: 10, color: { rgb: '15803D' } },
                fill:  { fgColor: { rgb: 'DCFCE7' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: bdr('15803D')
            };
            const sKpiLbl = (bg, fg) => ({
                font:      { bold: true, sz: 10, color: { rgb: fg } },
                fill:      { fgColor: { rgb: bg } },
                alignment: { horizontal: 'center', vertical: 'bottom' },
                border:    bdr(bg)
            });
            const sKpiVal = (bg, fg) => ({
                font:      { bold: true, sz: 28, color: { rgb: fg } },
                fill:      { fgColor: { rgb: bg } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border:    bdr(bg)
            });
            const sFooter = {
                font:      { italic: true, sz: 9, color: { rgb: '64748B' } },
                fill:      { fgColor: { rgb: 'F1F5F9' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border:    { top: { style: 'medium', color: { rgb: 'F59E0B' } } }
            };

            const buildSheet = (loans, accent, titulo) => {
                const ws     = XLS.utils.aoa_to_sheet([]);
                const COLS   = 10;
                ws['!cols']  = [
                    { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 16 },
                    { wch: 14 }, { wch: 12 }, { wch: 8 },
                    { wch: 22 }, { wch: 18 }, { wch: 30 }
                ];
                ws['!merges'] = [];
                ws['!rows']   = [];
                let row = 0;

                const sc  = (r, c, v, style, t = 's') => { ws[XLS.utils.encode_cell({ r, c })] = { v, t, s: style }; };
                const scN = (r, c, v, style)           => sc(r, c, v, style, 'n');
                const merge = (r1, c1, r2, c2)         => ws['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });

                // ── CABEÇALHO ─────────────────────────────────────────────────────
                ws['!rows'].push({ hpt: 46 });
                for (let c = 0; c < COLS; c++) sc(row, c, titulo, sTitle(accent));
                merge(row, 0, row, COLS - 1);
                row++;

                ws['!rows'].push({ hpt: 20 });
                const subtxt = `Emitido em ${today} às ${hora}  •  Total: ${loans.length} cautela(s)`;
                for (let c = 0; c < COLS; c++) sc(row, c, subtxt, sSub);
                merge(row, 0, row, COLS - 1);
                row++;

                // ── KPIs ──────────────────────────────────────────────────────────
                row++; // espaço
                ws['!rows'].push({ hpt: 6 }, { hpt: 30 }, { hpt: 34 }, { hpt: 18 });
                const openCnt     = loans.filter(l => l.status !== 'returned' && l.status !== 'damaged').length;
                const returnedCnt = loans.filter(l => l.status === 'returned').length;
                const damagedCnt  = loans.filter(l => l.status === 'damaged').length;
                const totalQty    = loans.reduce((s, l) => s + (l.quantity || 0), 0);
                const uniqueTools = new Set(loans.map(l => l.productName)).size;

                const cards = [
                    { bg: 'FEF9C3', fg: 'D97706', lbl: 'EM ABERTO',      val: openCnt },
                    { bg: 'DCFCE7', fg: '15803D', lbl: 'DEVOLVIDAS',     val: returnedCnt },
                    { bg: 'FFE4E6', fg: 'BE123C', lbl: 'DANIFICADAS',    val: damagedCnt },
                    { bg: 'F3E8FF', fg: '7E22CE', lbl: 'FERRAMENTAS',    val: uniqueTools }
                ];
                // Rótulos KPI (colunas 0,2,5,7 — agrupados em pares)
                const kpiCols = [[0, 1], [2, 3], [5, 6], [7, 8]];
                kpiCols.forEach(([c1, c2], i) => {
                    const { bg, fg, lbl, val } = cards[i];
                    for (let c = c1; c <= c2; c++) {
                        sc(row,   c, lbl, sKpiLbl(bg, fg));
                        scN(row+1, c, val, sKpiVal(bg, fg));
                    }
                    merge(row,   c1, row,   c2);
                    merge(row+1, c1, row+1, c2);
                });
                row += 2;
                row++; // espaço

                // ── SEÇÃO LABEL ───────────────────────────────────────────────────
                ws['!rows'].push({ hpt: 22 });
                for (let c = 0; c < COLS; c++) sc(row, c, `  REGISTRO DE CAUTELAS — ${titulo.toUpperCase()}`, sSecHdr(accent));
                merge(row, 0, row, COLS - 1);
                row++;

                // ── CABEÇALHO DA TABELA ───────────────────────────────────────────
                ws['!rows'].push({ hpt: 28 });
                const headers = ['STATUS', 'FERRAMENTA', 'DATA CAUTELA', 'DATA DEVOLUÇÃO',
                                 'CÓD. RM', 'CÓD. INTERNO', 'QTD', 'PESSOA', 'FUNÇÃO', 'OBSERVAÇÃO'];
                headers.forEach((h, c) => sc(row, c, h, sTH(accent)));
                row++;

                // ── LINHAS DE DADOS ───────────────────────────────────────────────
                const sorted = [...loans].sort((a, b) => (b.loanDate?.seconds || 0) - (a.loanDate?.seconds || 0));
                sorted.forEach((loan, idx) => {
                    const even      = idx % 2 === 0;
                    const isOpen    = loan.status !== 'returned' && loan.status !== 'damaged';
                    const isDamaged = loan.status === 'damaged';
                    ws['!rows'].push({ hpt: 20 });
                    const sBadgeStatus = isDamaged ? {
                        font:  { bold: true, sz: 10, color: { rgb: 'BE123C' } },
                        fill:  { fgColor: { rgb: 'FFE4E6' } },
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: bdr('BE123C')
                    } : isOpen ? sBadgeOpen : sBadgeDone;
                    const statusLabel = isDamaged ? 'DANIFICADA' : isOpen ? 'EM ABERTO' : 'DEVOLVIDA';

                    sc (row, 0, statusLabel, sBadgeStatus);
                    sc (row, 1, loan.productName || '', sCell(even, 'left', '1E293B', true));
                    sc (row, 2, formatFirestoreDate(loan.loanDate), sCell(even, 'center'));
                    sc (row, 3, isOpen ? '—' : formatFirestoreDate(loan.returnDate), sCell(even, 'center'));
                    sc (row, 4, loan.productCodeRM || '', sCell(even, 'center'));
                    sc (row, 5, loan.productCode || '', sCell(even, 'center'));
                    scN(row, 6, loan.quantity || 0, { ...sCell(even, 'center'), font: { bold: true, sz: 10, color: { rgb: '1D4ED8' } } });
                    sc (row, 7, loan.borrower || '', sCell(even, 'left', '1E293B', true));
                    sc (row, 8, loan.role || '', sCell(even));
                    sc (row, 9, loan.observation || '', sCell(even));
                    row++;
                });

                // ── RODAPÉ ────────────────────────────────────────────────────────
                row++;
                ws['!rows'].push({ hpt: 18 });
                const footerTxt = `EULLON  •  Ferramentaria  •  Gerado em ${today} ${hora}`;
                for (let c = 0; c < COLS; c++) sc(row, c, footerTxt, sFooter);
                merge(row, 0, row, COLS - 1);

                ws['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row, c: COLS - 1 } });
                return ws;
            };

            const openLoans     = toolLoans.filter(l => l.status !== 'returned' && l.status !== 'damaged');
            const returnedLoans = toolLoans.filter(l => l.status === 'returned');
            const damagedLoans  = toolLoans.filter(l => l.status === 'damaged');

            XLS.utils.book_append_sheet(wb, buildSheet(openLoans,     'F59E0B', 'Em Aberto'),  'Em Aberto');
            XLS.utils.book_append_sheet(wb, buildSheet(returnedLoans, '16A34A', 'Devolvidas'), 'Devolvidas');
            XLS.utils.book_append_sheet(wb, buildSheet(damagedLoans,  'DC2626', 'Danificadas'),'Danificadas');
            XLS.utils.book_append_sheet(wb, buildSheet([...toolLoans],'1D4ED8', 'Todas'),      'Todas');

            XLS.writeFile(wb, `cautelas_${stamp}.xlsx`);
            showToast('Relatório de cautelas exportado.');
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

                        const cells = columns.map(s => s.trim().replace(/"/g, ''));
                        const [codeRM, name, unitAbbr, quantity, location, minQuantity, group] = cells;
                        const unit = unitMap[unitAbbr.toLowerCase()] || unitAbbr;
                        const importPhotoRaw = cells.length >= 8 ? cells[7] : '';
                        const importPhotoUrl = sanitizeProductImageUrl(importPhotoRaw);

                        if (codeRM && name) {
                            const newProd = {
                                code: (nextCode++).toString(),
                                codeRM,
                                name: toUpperText(name),
                                unit,
                                location: toUpperText(location),
                                group: group || 'Outros',
                                quantity: parseInt(quantity) || 0,
                                minQuantity: parseInt(minQuantity) || 0,
                                observation: ''
                            };
                            if (importPhotoUrl) newProd.imageUrl = importPhotoUrl;
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
                                name: toUpperText(name),
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
            let csvContent = "Código RM;Nome;Unidade;Quantidade;Localização;Qtd. Mínima;Grupo;URL da foto\n";
            products.forEach(p => { 
                const row = [
                    p.codeRM || '',
                    p.name || '',
                    p.unit || '',
                    p.quantity || 0,
                    p.location || '',
                    p.minQuantity || 0,
                    p.group || '',
                    p.imageUrl || ''
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

            let csvContent = "Data;Produto;Código RM;Cód. Interno;Grupo;Quantidade;Retirado por;Função do Funcionário;Obra;Local de Aplicação;Código do Local;Detalhes\n";
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
            const XLS = typeof XLSX !== 'undefined' ? XLSX : null;
            if (!XLS) { showToast('Biblioteca de planilhas não carregou. Recarregue a página.', true); return; }

            try {
                // ── DADOS ──────────────────────────────────────────────────────
                const totalProducts = products.length;
                const totalUnits    = products.reduce((s, p) => s + (p.quantity || 0), 0);
                const lowStockItems = products.filter(p => (p.quantity || 0) <= (p.minQuantity || 0));
                const lowStockCount = lowStockItems.length;
                const zeroCount     = products.filter(p => (p.quantity || 0) === 0).length;
                const pendingReqs   = requisitions.filter(r => r.status === 'Pendente' || r.status === 'pending').length;

                const tsMs = (e) => {
                    if (!e?.date) return 0;
                    if (typeof e.date.seconds === 'number') return e.date.seconds * 1000;
                    if (typeof e.date.toMillis === 'function') return e.date.toMillis();
                    return 0;
                };
                const getTopExits = (days) => {
                    const cutoff = Date.now() - days * 864e5;
                    const map = new Map();
                    history.filter(h => (h.type === 'Saída' || h.type === 'Saída por Requisição') && tsMs(h) >= cutoff)
                        .forEach(h => {
                            const k = h.productId || h.productName;
                            const v = map.get(k) || { name: h.productName || '—', code: h.productCodeRM || h.productCode || '—', qty: 0 };
                            v.qty += Math.abs(Number(h.quantity) || 0);
                            map.set(k, v);
                        });
                    return [...map.values()].sort((a, b) => b.qty - a.qty).slice(0, 10);
                };
                const entQty = (days) => {
                    const c = Date.now() - days * 864e5;
                    return history.filter(h => (h.type === 'Entrada' || h.type === 'Ajuste Entrada') && tsMs(h) >= c)
                        .reduce((s, h) => s + Math.abs(Number(h.quantity) || 0), 0);
                };
                const saiQty = (days) => {
                    const c = Date.now() - days * 864e5;
                    return history.filter(h => (h.type === 'Saída' || h.type === 'Saída por Requisição') && tsMs(h) >= c)
                        .reduce((s, h) => s + Math.abs(Number(h.quantity) || 0), 0);
                };

                const top3 = getTopExits(3), top7 = getTopExits(7), top30 = getTopExits(30);
                const e7 = entQty(7), s7 = saiQty(7), e30 = entQty(30), s30 = saiQty(30);

                const groupMap = new Map();
                products.forEach(p => {
                    const g = p.group || 'Sem Grupo';
                    const v = groupMap.get(g) || { count: 0, units: 0, low: 0, zero: 0 };
                    v.count++;
                    v.units += p.quantity || 0;
                    if ((p.quantity || 0) === 0) v.zero++;
                    else if ((p.quantity || 0) <= (p.minQuantity || 0)) v.low++;
                    groupMap.set(g, v);
                });
                const groupData = [...groupMap.entries()].sort((a, b) => b[1].units - a[1].units);

                const today     = new Date().toLocaleDateString('pt-BR');
                const hora      = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
                const timestamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
                const wb        = XLS.utils.book_new();

                // ── ESTILOS ────────────────────────────────────────────────────
                const bdr = (color = 'E2E8F0') => ({
                    top:    { style: 'thin', color: { rgb: color } },
                    bottom: { style: 'thin', color: { rgb: color } },
                    left:   { style: 'thin', color: { rgb: color } },
                    right:  { style: 'thin', color: { rgb: color } }
                });

                // Títulos de seção (faixa escura)
                const sTitle = (accentColor = '0066FF') => ({
                    font:      { bold: true, sz: 20, color: { rgb: 'FFFFFF' } },
                    fill:      { fgColor: { rgb: '0F172A' } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border:    { bottom: { style: 'thick', color: { rgb: accentColor } } }
                });
                const sSub = {
                    font:      { italic: true, sz: 10, color: { rgb: '94A3B8' } },
                    fill:      { fgColor: { rgb: '1E293B' } },
                    alignment: { horizontal: 'center', vertical: 'center' }
                };
                const sSecHdr = (rgb) => ({
                    font:      { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
                    fill:      { fgColor: { rgb } },
                    alignment: { horizontal: 'left', vertical: 'center' },
                    border:    bdr(rgb)
                });
                const sTH = (rgb = '1E40AF') => ({
                    font:      { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
                    fill:      { fgColor: { rgb } },
                    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                    border:    bdr(rgb)
                });
                const sCell = (even, align = 'left', fgColor = '111827', bold = false) => ({
                    font:      { bold, sz: 10, color: { rgb: fgColor } },
                    fill:      { fgColor: { rgb: even ? 'F8FAFC' : 'FFFFFF' } },
                    alignment: { horizontal: align, vertical: 'center' },
                    border:    bdr()
                });
                const sFooter = {
                    font:      { italic: true, sz: 9, color: { rgb: '64748B' } },
                    fill:      { fgColor: { rgb: 'F1F5F9' } },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border:    { top: { style: 'medium', color: { rgb: '0066FF' } } }
                };

                // Cards de KPI: (bgColor, textColor, valueColor)
                const kpiCard = (bg, text, val) => ({
                    lbl: { font: { bold: true, sz: 10, color: { rgb: text } }, fill: { fgColor: { rgb: bg } }, alignment: { horizontal: 'center', vertical: 'bottom' }, border: bdr(bg) },
                    val: { font: { bold: true, sz: 26, color: { rgb: val } }, fill: { fgColor: { rgb: bg } }, alignment: { horizontal: 'center', vertical: 'center' }, border: bdr(bg) },
                    sub: { font: { sz: 9, italic: true, color: { rgb: text } }, fill: { fgColor: { rgb: bg } }, alignment: { horizontal: 'center', vertical: 'top' }, border: bdr(bg) }
                });
                const cBlue  = kpiCard('DBEAFE', '1D4ED8', '1D4ED8');
                const cGreen = kpiCard('DCFCE7', '15803D', '15803D');
                const cAmber = kpiCard('FEF9C3', 'A16207', 'D97706');
                const cRed   = kpiCard('FFE4E6', 'BE123C', 'DC2626');

                // ────────────────────────────────────────────────────────────────
                // ABA 1 — PAINEL DE KPIs  (8 colunas A..H)
                // ────────────────────────────────────────────────────────────────
                const ws = XLS.utils.aoa_to_sheet([]);
                ws['!cols']   = [{ wch: 22 }, { wch: 14 }, { wch: 22 }, { wch: 14 }, { wch: 22 }, { wch: 14 }, { wch: 22 }, { wch: 14 }];
                ws['!merges'] = [];
                ws['!rows']   = [];

                let row = 0;
                const sc  = (r, c, v, style, t = 's') => { ws[XLS.utils.encode_cell({ r, c })] = { v, t, s: style }; };
                const scN = (r, c, v, style)           => sc(r, c, v, style, 'n');
                const merge = (r1, c1, r2, c2)         => ws['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
                const hpt   = (r, h)                   => { ws['!rows'][r] = { hpt: h }; };

                // Cabeçalho principal
                sc(row, 0, 'RELATÓRIO DE KPIs — ESTOQUE UHE ESTRELA', sTitle()); merge(row, 0, row, 7); hpt(row, 44); row++;
                sc(row, 0, `Gerado em ${today} às ${hora}  ·  Dados em tempo real`, sSub); merge(row, 0, row, 7); hpt(row, 22); row++;
                row++; // espaço

                // ── Cards de KPI (3 linhas por card: label / valor / sublabel) ──
                const cards = [
                    { lbl: 'TOTAL DE PRODUTOS', val: totalProducts,  sub: 'SKUs cadastrados',        card: cBlue  },
                    { lbl: 'TOTAL DE UNIDADES',  val: totalUnits,    sub: 'Unidades em estoque',      card: cGreen },
                    { lbl: 'ESTOQUE BAIXO',       val: lowStockCount, sub: 'Itens abaixo do mínimo',  card: cAmber },
                    { lbl: 'REQ. PENDENTES',      val: pendingReqs,   sub: 'Aguardando atendimento',  card: cRed   },
                ];
                hpt(row, 24); hpt(row + 1, 44); hpt(row + 2, 20);
                cards.forEach(({ lbl, val, sub, card }, i) => {
                    const c = i * 2;
                    sc(row, c, lbl, card.lbl); sc(row, c + 1, lbl, card.lbl); merge(row, c, row, c + 1);
                    scN(row + 1, c, val, card.val); scN(row + 1, c + 1, val, card.val); merge(row + 1, c, row + 1, c + 1);
                    sc(row + 2, c, sub, card.sub); sc(row + 2, c + 1, sub, card.sub); merge(row + 2, c, row + 2, c + 1);
                });
                row += 4; // 3 linhas card + 1 espaço

                // ── Movimentação ──
                sc(row, 0, '  MOVIMENTAÇÃO DE ESTOQUE', sSecHdr('1E3A5F')); merge(row, 0, row, 7); hpt(row, 26); row++;
                ['Período', 'Entradas (Un.)', 'Saídas (Un.)', 'Saldo', '', '', '', ''].forEach((h, c) => {
                    if (c < 4) sc(row, c, h, sTH('1E40AF'));
                });
                merge(row, 4, row, 7); hpt(row, 22); row++;

                [{ lbl: '7 dias', e: e7, s: s7 }, { lbl: '30 dias', e: e30, s: s30 }].forEach(({ lbl, e, s }, i) => {
                    const ev = i % 2 === 1, saldo = e - s;
                    sc(row, 0, `Últimos ${lbl}`, sCell(ev, 'left', '111827', true));
                    scN(row, 1, e, sCell(ev, 'center', '15803D', true));
                    scN(row, 2, s, sCell(ev, 'center', 'B91C1C', true));
                    scN(row, 3, saldo, sCell(ev, 'center', saldo >= 0 ? '15803D' : 'DC2626', true));
                    merge(row, 4, row, 7); sc(row, 4, '', sCell(ev)); hpt(row, 20); row++;
                });
                row++; // espaço

                // ── Distribuição por Grupo ──
                sc(row, 0, '  DISTRIBUIÇÃO POR GRUPO DE MATERIAL', sSecHdr('312E81')); merge(row, 0, row, 7); hpt(row, 26); row++;
                ['Grupo', 'SKUs', 'Unidades', 'Est. Baixo', 'Zerados', 'OK', '', ''].forEach((h, c) => {
                    if (c < 6) sc(row, c, h, sTH('4338CA'));
                });
                merge(row, 6, row, 7); hpt(row, 22); row++;

                groupData.forEach(([g, d], i) => {
                    const ev = i % 2 === 1, ok = d.count - d.zero - d.low;
                    sc(row, 0, g, sCell(ev, 'left', '111827', true));
                    scN(row, 1, d.count, sCell(ev, 'center'));
                    scN(row, 2, d.units, sCell(ev, 'center'));
                    scN(row, 3, d.low,  sCell(ev, 'center', d.low  > 0 ? 'D97706' : '111827', d.low  > 0));
                    scN(row, 4, d.zero, sCell(ev, 'center', d.zero > 0 ? 'DC2626' : '111827', d.zero > 0));
                    scN(row, 5, ok,     sCell(ev, 'center', '15803D'));
                    merge(row, 6, row, 7); sc(row, 6, '', sCell(ev)); hpt(row, 18); row++;
                });
                row++;

                // ── Top Saídas ──
                const renderTop = (title, items, secRgb, thRgb) => {
                    sc(row, 0, `  ${title}`, sSecHdr(secRgb)); merge(row, 0, row, 7); hpt(row, 26); row++;
                    ['#', 'Material', 'Código RM', 'Qtd. Saída', '', '', '', ''].forEach((h, c) => {
                        if (c < 4) sc(row, c, h, sTH(thRgb));
                    });
                    merge(row, 4, row, 7); hpt(row, 22); row++;
                    if (!items.length) {
                        sc(row, 0, 'Sem registros no período', sCell(false, 'center', '94A3B8'));
                        merge(row, 0, row, 7); hpt(row, 18); row++;
                    } else {
                        items.forEach((item, i) => {
                            const ev    = i % 2 === 1;
                            const medal = ['🥇', '🥈', '🥉'][i] || `${i + 1}º`;
                            sc(row, 0, medal,     sCell(ev, 'center', '111827', true));
                            sc(row, 1, item.name, sCell(ev, 'left'));
                            sc(row, 2, item.code, sCell(ev, 'center'));
                            scN(row, 3, item.qty, sCell(ev, 'center', 'DC2626', true));
                            merge(row, 4, row, 7); sc(row, 4, '', sCell(ev)); hpt(row, 18); row++;
                        });
                    }
                    row++;
                };

                renderTop('TOP SAÍDAS — ÚLTIMOS 3 DIAS',  top3,  '7F1D1D', 'B91C1C');
                renderTop('TOP SAÍDAS — ÚLTIMOS 7 DIAS',  top7,  '7C2D12', 'C2410C');
                renderTop('TOP SAÍDAS — ÚLTIMOS 30 DIAS', top30, '78350F', 'D97706');

                // Rodapé
                sc(row, 0, `UHE Estrela  ·  KPI Report  ·  ${today} às ${hora}`, sFooter);
                merge(row, 0, row, 7); hpt(row, 18); row++;

                ws['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row - 1, c: 7 } });
                XLS.utils.book_append_sheet(wb, ws, 'Painel de KPIs');

                // ────────────────────────────────────────────────────────────────
                // ABA 2 — ESTOQUE BAIXO (detalhado, todos os itens)
                // ────────────────────────────────────────────────────────────────
                const lowSorted = [...lowStockItems].sort((a, b) => (a.quantity - a.minQuantity) - (b.quantity - b.minQuantity));

                const ws2     = XLS.utils.aoa_to_sheet([]);
                ws2['!cols']  = [{ wch: 5 }, { wch: 38 }, { wch: 14 }, { wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }];
                ws2['!merges'] = [];
                ws2['!rows']   = [];

                let r2 = 0;
                const sc2  = (r, c, v, style, t = 's') => { ws2[XLS.utils.encode_cell({ r, c })] = { v, t, s: style }; };
                const scN2 = (r, c, v, style)           => sc2(r, c, v, style, 'n');
                const mg2  = (r1, c1, r2e, c2)          => ws2['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2e, c: c2 } });
                const hp2  = (r, h)                      => { ws2['!rows'][r] = { hpt: h }; };

                sc2(r2, 0, `ITENS COM ESTOQUE BAIXO — UHE ESTRELA  (${lowSorted.length} itens)`, sTitle('DC2626')); mg2(r2, 0, r2, 7); hp2(r2, 44); r2++;
                sc2(r2, 0, `Gerado em ${today} às ${hora}  ·  Ordenado por maior déficit`, sSub); mg2(r2, 0, r2, 7); hp2(r2, 22); r2++;
                r2++;

                ['Nº', 'Material', 'Código RM', 'Grupo', 'Estoque', 'Mínimo', 'Déficit', 'Status'].forEach((h, c) => {
                    sc2(r2, c, h, sTH('991B1B'));
                });
                hp2(r2, 24); r2++;

                if (!lowSorted.length) {
                    sc2(r2, 0, '✅  Nenhum item com estoque baixo — tudo dentro dos limites!', {
                        font: { bold: true, sz: 12, color: { rgb: '15803D' } },
                        fill: { fgColor: { rgb: 'DCFCE7' } },
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: bdr('15803D')
                    });
                    mg2(r2, 0, r2, 7); hp2(r2, 30); r2++;
                } else {
                    lowSorted.forEach((p, i) => {
                        const ev     = i % 2 === 1;
                        const qty    = p.quantity || 0;
                        const min    = p.minQuantity || 0;
                        const deficit = min - qty;
                        const status = qty === 0 ? '🔴 ZERADO' : deficit >= min * 0.5 ? '🟠 CRÍTICO' : '🟡 BAIXO';
                        const statusC = qty === 0 ? 'B91C1C' : deficit >= min * 0.5 ? 'C2410C' : 'D97706';

                        scN2(r2, 0, i + 1,                        sCell(ev, 'center'));
                        sc2(r2,  1, p.name || '—',                sCell(ev, 'left'));
                        sc2(r2,  2, p.codeRM || p.code || '—',    sCell(ev, 'center'));
                        sc2(r2,  3, p.group || '—',               sCell(ev, 'center'));
                        scN2(r2, 4, qty,                          sCell(ev, 'center', qty === 0 ? 'DC2626' : '111827', qty === 0));
                        scN2(r2, 5, min,                          sCell(ev, 'center'));
                        scN2(r2, 6, deficit > 0 ? -deficit : 0,  sCell(ev, 'center', deficit > 0 ? 'DC2626' : '15803D', true));
                        sc2(r2,  7, status,                       sCell(ev, 'center', statusC, true));
                        hp2(r2, 18); r2++;
                    });
                }

                sc2(r2, 0, `UHE Estrela  ·  Estoque Baixo  ·  ${today}`, sFooter);
                mg2(r2, 0, r2, 7); hp2(r2, 18); r2++;

                ws2['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r2 - 1, c: 7 } });
                XLS.utils.book_append_sheet(wb, ws2, 'Estoque Baixo');

                // ── DOWNLOAD ───────────────────────────────────────────────────
                XLS.writeFile(wb, `kpis_estoque_${timestamp}.xlsx`);
                showToast('✅ Relatório de KPIs exportado com sucesso!');

            } catch (err) {
                console.error('Erro ao gerar KPI Excel:', err);
                showToast(`Erro ao gerar relatório: ${err.message}`, true);
            }
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

            // Exportar Sugestão de Compras — Excel .xlsx (várias abas)
            if (e.target.id === 'compras-excel-btn' || e.target.closest('#compras-excel-btn')) {
                const XLS = typeof XLSX !== 'undefined' ? XLSX : null;
                if (!XLS) {
                    showToast('Biblioteca Excel indisponível. Recarregue a página.', true);
                    return;
                }

                const periodLabel = (() => {
                    const sel = document.getElementById('compras-period-select');
                    return sel ? sel.options[sel.selectedIndex].text : `Últimos ${lastComprasPeriodDays} dias`;
                })();
                const covSel = document.getElementById('compras-coverage-select');
                const coverageLabel = covSel ? covSel.options[covSel.selectedIndex].text : `Cobrir ${lastComprasExport.coverageDays} dias`;

                const ex = lastComprasExport;
                const nPed = ex.pedido?.length || 0;
                const nTor = ex.turnover?.length || 0;
                const nZer = ex.zero?.length || 0;
                const nLow = ex.low?.length || 0;

                if (!nPed && !nTor && !nZer && !nLow) {
                    showToast('Abra a tela Sugestão de Compras para carregar os dados antes de exportar.', true);
                    return;
                }

                const urgencyLabel = (s) => (s === 'critico' ? 'CRÍTICO' : s === 'atencao' ? 'ATENÇÃO' : 'MONITORAR');
                const stamp = new Date().toISOString().slice(0, 10);
                const agora = new Date().toLocaleString('pt-BR');

                try {
                    const wb = XLS.utils.book_new();

                    const bdr = (color = 'CBD5E1') => ({
                        top: { style: 'thin', color: { rgb: color } },
                        bottom: { style: 'thin', color: { rgb: color } },
                        left: { style: 'thin', color: { rgb: color } },
                        right: { style: 'thin', color: { rgb: color } }
                    });
                    const sTitle = {
                        font: { bold: true, sz: 16, color: { rgb: 'FFFFFF' } },
                        fill: { fgColor: { rgb: '0F172A' } },
                        alignment: { horizontal: 'center', vertical: 'center' },
                        border: bdr('1E3A8A')
                    };
                    const sSub = {
                        font: { italic: true, sz: 10, color: { rgb: '64748B' } },
                        fill: { fgColor: { rgb: 'F1F5F9' } },
                        alignment: { horizontal: 'center', vertical: 'center' }
                    };
                    const sTH = {
                        font: { bold: true, sz: 10, color: { rgb: 'FFFFFF' } },
                        fill: { fgColor: { rgb: '1E40AF' } },
                        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                        border: bdr('1E40AF')
                    };
                    const sCell = (even) => ({
                        font: { sz: 10, color: { rgb: '111827' } },
                        fill: { fgColor: { rgb: even ? 'F8FAFC' : 'FFFFFF' } },
                        alignment: { vertical: 'center' },
                        border: bdr()
                    });
                    const sNum = (even, align = 'right') => ({
                        font: { sz: 10, color: { rgb: '111827' } },
                        fill: { fgColor: { rgb: even ? 'F8FAFC' : 'FFFFFF' } },
                        alignment: { horizontal: align, vertical: 'center' },
                        border: bdr()
                    });

                    const mergeRange = (ws, r1, c1, r2, c2) => {
                        if (!ws['!merges']) ws['!merges'] = [];
                        ws['!merges'].push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
                    };
                    const setCell = (ws, r, c, v, style, t = 's') => {
                        ws[XLS.utils.encode_cell({ r, c })] = { v, t, s: style };
                    };
                    const setNum = (ws, r, c, v, style) => setCell(ws, r, c, v, style, typeof v === 'number' ? 'n' : 's');

                    // ── Aba Resumo ─────────────────────────────────────────────
                    const wsRes = XLS.utils.aoa_to_sheet([]);
                    wsRes['!cols'] = [{ wch: 52 }, { wch: 24 }];
                    let r0 = 0;
                    setCell(wsRes, r0, 0, 'SUGESTÃO DE COMPRAS — RESUMO', sTitle);
                    mergeRange(wsRes, r0, 0, r0, 1);
                    r0++;
                    setCell(wsRes, r0, 0, `${periodLabel} · ${coverageLabel}`, sSub);
                    mergeRange(wsRes, r0, 0, r0, 1);
                    r0++;
                    setCell(wsRes, r0, 0, `Gerado em: ${agora}`, sSub);
                    mergeRange(wsRes, r0, 0, r0, 1);
                    r0 += 2;

                    const totalPedQtd = (ex.pedido || []).reduce((a, i) => a + (i.suggestedQty || 0), 0);
                    const metaStyle = sCell(false);
                    metaStyle.alignment = { vertical: 'center', wrapText: true };
                    const linhas = [
                        `Itens com quantidade sugerida para compra (aba "Pedido sugerido"): ${nPed}`,
                        `Quantidade total sugerida (somatório): ${totalPedQtd}`,
                        `Itens com maior saída no período (aba "Rotatividade"): ${nTor}`,
                        `Produtos com estoque zerado (aba "Estoque zerado"): ${nZer}`,
                        `Produtos abaixo do estoque mínimo cadastrado (aba "Abaixo mínimo"): ${nLow}`,
                        '',
                        'Use a aba "Pedido sugerido" para montar o pedido. As demais abas ajudam a priorizar por giro e risco de ruptura.'
                    ];
                    linhas.forEach((txt, i) => {
                        setCell(wsRes, r0 + i, 0, txt, metaStyle);
                    });
                    r0 += linhas.length;
                    wsRes['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r0, c: 1 } });
                    XLS.utils.book_append_sheet(wb, wsRes, 'Resumo');

                    const appendTableSheet = (sheetName, headers, rowCells) => {
                        const ws = XLS.utils.aoa_to_sheet([]);
                        const colW = headers.map((h) => ({ wch: Math.min(40, Math.max(10, String(h).length + 5)) }));
                        ws['!cols'] = colW;
                        let rr = 0;
                        setCell(ws, rr, 0, sheetName.toUpperCase(), sTitle);
                        mergeRange(ws, rr, 0, rr, headers.length - 1);
                        rr++;
                        setCell(ws, rr, 0, `${periodLabel} · ${coverageLabel} · ${agora}`, sSub);
                        mergeRange(ws, rr, 0, rr, headers.length - 1);
                        rr += 2;

                        headers.forEach((h, c) => setCell(ws, rr, c, h, sTH));
                        rr++;

                        rowCells.forEach((cells, idx) => {
                            const even = idx % 2 === 0;
                            cells.forEach((cell, c) => {
                                const v = cell.v;
                                const isNum = cell.n;
                                const align = cell.a || 'left';
                                const st = isNum ? sNum(even, align === 'right' ? 'right' : 'center') : { ...sCell(even), alignment: { horizontal: align, vertical: 'center' } };
                                if (isNum) setNum(ws, rr, c, v, st);
                                else setCell(ws, rr, c, v ?? '—', st);
                            });
                            rr++;
                        });

                        ws['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rr - 1, c: headers.length - 1 } });
                        const safeName = sheetName.replace(/[:\\/?*[\]]/g, '').slice(0, 31);
                        XLS.utils.book_append_sheet(wb, ws, safeName);
                    };

                    // ── Pedido sugerido ───────────────────────────────────────
                    if (nPed) {
                        const headers = ['Urgência', 'Material', 'Código RM', 'Grupo', 'Unid.', 'Local', 'Estoque', 'Mín.', 'Saída período', 'Média/dia', 'Dias rest.', 'Qtd compra', 'Última saída'];
                        const rowCells = ex.pedido.map((i) => {
                            const { p, dailyRate, daysRemaining, suggestedQty, status, lastMov, totalConsumed } = i;
                            const lastStr = lastMov ? new Date(lastMov).toLocaleDateString('pt-BR') : '—';
                            return [
                                { v: urgencyLabel(status), a: 'center' },
                                { v: p.name || '' },
                                { v: p.codeRM || p.code || '—' },
                                { v: p.group || '—' },
                                { v: p.unit || 'UN', a: 'center' },
                                { v: p.location || '—' },
                                { v: Number(p.quantity ?? 0), n: true, a: 'right' },
                                { v: Number(p.minQuantity ?? 0), n: true, a: 'right' },
                                { v: Number(totalConsumed || 0), n: true, a: 'right' },
                                { v: dailyRate > 0 ? Number(dailyRate.toFixed(3)) : 0, n: true, a: 'right' },
                                { v: daysRemaining !== null && daysRemaining !== undefined ? Number(daysRemaining) : '—', n: daysRemaining !== null && daysRemaining !== undefined, a: 'right' },
                                { v: Number(suggestedQty), n: true, a: 'right' },
                                { v: lastStr, a: 'center' }
                            ];
                        });
                        appendTableSheet('Pedido sugerido', headers, rowCells);
                    }

                    // ── Rotatividade ──────────────────────────────────────────
                    if (nTor) {
                        const headers = ['Pos.', 'Material', 'Código RM', 'Grupo', 'Unid.', 'Local', 'Saída período', 'Média/dia', 'Estoque', 'Qtd sugerida'];
                        const rowCells = ex.turnover.map((row, idx) => {
                            const { p, totalConsumed, dailyRate, qty, suggestedQty } = row;
                            return [
                                { v: idx + 1, n: true, a: 'center' },
                                { v: p.name || '' },
                                { v: p.codeRM || p.code || '—' },
                                { v: p.group || '—' },
                                { v: p.unit || 'UN', a: 'center' },
                                { v: p.location || '—' },
                                { v: Number(totalConsumed), n: true, a: 'right' },
                                { v: dailyRate > 0 ? Number(dailyRate.toFixed(3)) : 0, n: true, a: 'right' },
                                { v: Number(qty), n: true, a: 'right' },
                                { v: suggestedQty > 0 ? Number(suggestedQty) : '—', n: suggestedQty > 0, a: 'right' }
                            ];
                        });
                        appendTableSheet('Rotatividade', headers, rowCells);
                    }

                    // ── Estoque zerado ─────────────────────────────────────────
                    if (nZer) {
                        const headers = ['Material', 'Código RM', 'Grupo', 'Unid.', 'Local', 'Mín.', 'Saída período', 'Média/dia', 'Qtd sugerida'];
                        const rowCells = ex.zero.map((row) => {
                            const { p, min, totalConsumed, dailyRate, suggestedQty } = row;
                            return [
                                { v: p.name || '' },
                                { v: p.codeRM || p.code || '—' },
                                { v: p.group || '—' },
                                { v: p.unit || 'UN', a: 'center' },
                                { v: p.location || '—' },
                                { v: Number(min), n: true, a: 'right' },
                                { v: Number(totalConsumed), n: true, a: 'right' },
                                { v: dailyRate > 0 ? Number(dailyRate.toFixed(3)) : 0, n: true, a: 'right' },
                                { v: suggestedQty > 0 ? Number(suggestedQty) : '—', n: suggestedQty > 0, a: 'right' }
                            ];
                        });
                        appendTableSheet('Estoque zerado', headers, rowCells);
                    }

                    // ── Abaixo do mínimo ──────────────────────────────────────
                    if (nLow) {
                        const headers = ['Material', 'Código RM', 'Grupo', 'Unid.', 'Local', 'Estoque', 'Mín.', 'Saída período', 'Média/dia', 'Qtd sugerida'];
                        const rowCells = ex.low.map((row) => {
                            const { p, qty, min, totalConsumed, dailyRate, suggestedQty } = row;
                            return [
                                { v: p.name || '' },
                                { v: p.codeRM || p.code || '—' },
                                { v: p.group || '—' },
                                { v: p.unit || 'UN', a: 'center' },
                                { v: p.location || '—' },
                                { v: Number(qty), n: true, a: 'right' },
                                { v: Number(min), n: true, a: 'right' },
                                { v: Number(totalConsumed), n: true, a: 'right' },
                                { v: dailyRate > 0 ? Number(dailyRate.toFixed(3)) : 0, n: true, a: 'right' },
                                { v: suggestedQty > 0 ? Number(suggestedQty) : '—', n: suggestedQty > 0, a: 'right' }
                            ];
                        });
                        appendTableSheet('Abaixo mínimo', headers, rowCells);
                    }

                    XLS.writeFile(wb, `compras_pedido_${lastComprasPeriodDays}d_${stamp}.xlsx`);
                    showToast('Planilha Excel (.xlsx) exportada com resumo, pedido, giro e rupturas.');
                } catch (err) {
                    console.error('Export compras Excel:', err);
                    showToast(`Erro ao exportar: ${err.message}`, true);
                }
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
            if (e.target.id === 'compras-period-select' ||
                e.target.id === 'compras-coverage-select' ||
                e.target.id === 'compras-group-filter') {
                const sel = document.getElementById('compras-period-select');
                renderComprasView(sel ? parseInt(sel.value) : 15);
            }
        });

        // Busca por nome em tempo real (#6)
        document.addEventListener('input', (e) => {
            if (e.target.id === 'compras-search') {
                const sel = document.getElementById('compras-period-select');
                renderComprasView(sel ? parseInt(sel.value) : 15);
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

        searchInput.addEventListener('input', () => {
            currentPage = 1; // 🚀 Resetar para primeira página ao buscar
            renderProducts();
        });
        exitsSearchInput.addEventListener('input', (e) => renderExitLog(e.target.value));
        activityLogSearchInput?.addEventListener('input', (e) => renderActivityLog(e.target.value));
        rmSearchInput.addEventListener('input', (e) => renderRMView(e.target.value));
        toolLoanSearchInput?.addEventListener('input', populateToolLoanProducts);
        
        backupBtn.addEventListener('click', async () => {
             try {
                showLoader(true);
                const backupData = {
                    products,
                    history,
                    locations,
                    requisitions,
                    toolLoans,
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
                                    locationsCollectionRef, requisitionsCollectionRef,
                                    toolLoansCollectionRef
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
                                    requisitions: requisitionsCollectionRef,
                                    toolLoans: toolLoansCollectionRef
                                };
                                
                                for(const key in collectionsToRestore) {
                                    if(backupData[key] && Array.isArray(backupData[key])) {
                                        restoreTimestamps(backupData[key]).forEach(item => {
                                            const docRef = item.id ? doc(collectionsToRestore[key], item.id) : doc(collectionsToRestore[key]);
                                            const { id, ...data } = item;
                                            restoreBatch.set(docRef, normalizeRecordForUppercase(key, data));
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
                    const name = toUpperText(document.getElementById('ep-name').value);
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
                            location: toUpperText(document.getElementById('ep-location').value),
                            createdAt: serverTimestamp(),
                            createdBy: toUpperText(currentUser?.displayName || 'Anônimo')
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
                        name: toUpperText(document.getElementById('epe-name').value),
                        codeRM: document.getElementById('epe-code-rm').value.trim(),
                        group: document.getElementById('epe-group').value,
                        unit: document.getElementById('epe-unit').value,
                        quantity: parseInt(document.getElementById('epe-qty').value) || 0,
                        minQuantity: parseInt(document.getElementById('epe-min').value) || 0,
                        location: toUpperText(document.getElementById('epe-location').value),
                        updatedAt: serverTimestamp(),
                        updatedBy: toUpperText(currentUser?.displayName || 'Anônimo')
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
                const supplier = toUpperText(document.getElementById('estrela-entry-supplier').value);
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
                        receivedBy: toUpperOrNull(document.getElementById('estrela-entry-received-by').value),
                        observation: toUpperOrNull(document.getElementById('estrela-entry-obs').value),
                        date: serverTimestamp(),
                        registeredBy: toUpperText(currentUser?.displayName || 'Anônimo')
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
                const who = toUpperText(document.getElementById('estrela-exit-who').value);
                const leader = toUpperText(document.getElementById('estrela-exit-leader').value);
                const appLocation = toUpperText(document.getElementById('estrela-exit-location').value);

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
                        observation: toUpperOrNull(document.getElementById('estrela-exit-obs').value),
                        date: serverTimestamp(),
                        registeredBy: toUpperText(currentUser?.displayName || 'Anônimo')
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
            try {

            const wb = XLS.utils.book_new();
            const today = new Date().toLocaleDateString('pt-BR');
            const hora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

            // ─── ESTILOS PROFISSIONAIS ───
            const titleStyle = {
                font: { bold: true, sz: 18, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '1F2937' } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: { bottom: { style: 'medium', color: { rgb: '0066FF' } } }
            };
            const subtitleStyle = {
                font: { italic: true, sz: 11, color: { rgb: '374151' } },
                alignment: { horizontal: 'center', vertical: 'center' }
            };
            const headerStyle = {
                font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
                fill: { fgColor: { rgb: '0066FF' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                    top: { style: 'thin', color: { rgb: '003D8A' } },
                    bottom: { style: 'thin', color: { rgb: '003D8A' } },
                    left: { style: 'thin', color: { rgb: '003D8A' } },
                    right: { style: 'thin', color: { rgb: '003D8A' } }
                }
            };
            const headerGreen = { ...headerStyle, fill: { fgColor: { rgb: '059669' } }, border: { ...headerStyle.border, top: { style: 'thin', color: { rgb: '047857' } }, bottom: { style: 'thin', color: { rgb: '047857' } }, left: { style: 'thin', color: { rgb: '047857' } }, right: { style: 'thin', color: { rgb: '047857' } } } };
            const headerRed = { ...headerStyle, fill: { fgColor: { rgb: 'DC2626' } }, border: { ...headerStyle.border, top: { style: 'thin', color: { rgb: 'B91C1C' } }, bottom: { style: 'thin', color: { rgb: 'B91C1C' } }, left: { style: 'thin', color: { rgb: 'B91C1C' } }, right: { style: 'thin', color: { rgb: 'B91C1C' } } } };
            const headerPurple = { ...headerStyle, fill: { fgColor: { rgb: '7C3AED' } }, border: { ...headerStyle.border, top: { style: 'thin', color: { rgb: '6D28D9' } }, bottom: { style: 'thin', color: { rgb: '6D28D9' } }, left: { style: 'thin', color: { rgb: '6D28D9' } }, right: { style: 'thin', color: { rgb: '6D28D9' } } } };

            const cellBorder = {
                top: { style: 'thin', color: { rgb: 'E5E7EB' } },
                bottom: { style: 'thin', color: { rgb: 'E5E7EB' } },
                left: { style: 'thin', color: { rgb: 'E5E7EB' } },
                right: { style: 'thin', color: { rgb: 'E5E7EB' } }
            };
            const cellBase = { font: { sz: 10, color: { rgb: '111827' } }, alignment: { vertical: 'center' }, border: cellBorder };
            const cellCenter = { ...cellBase, alignment: { horizontal: 'center', vertical: 'center' } };
            const cellBold = { ...cellCenter, font: { sz: 10, bold: true, color: { rgb: '111827' } } };
            const cellEven = { ...cellBase, fill: { fgColor: { rgb: 'F9FAFB' } } };
            const cellEvenCenter = { ...cellCenter, fill: { fgColor: { rgb: 'F9FAFB' } } };
            const cellEvenBold = { ...cellBold, fill: { fgColor: { rgb: 'F9FAFB' } } };
            const cellEvenGreen = { ...cellBase, fill: { fgColor: { rgb: 'F0FDF4' } } };
            const cellEvenGreenCenter = { ...cellCenter, fill: { fgColor: { rgb: 'F0FDF4' } } };
            const cellEvenRed = { ...cellBase, fill: { fgColor: { rgb: 'FEF2F2' } } };
            const cellEvenRedCenter = { ...cellCenter, fill: { fgColor: { rgb: 'FEF2F2' } } };
            const cellEvenPurple = { ...cellBase, fill: { fgColor: { rgb: 'FAF5FF' } } };
            const cellEvenPurpleCenter = { ...cellCenter, fill: { fgColor: { rgb: 'FAF5FF' } } };

            // Status com cores mais vibrantes
            const statusOk = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: '065F46' } }, fill: { fgColor: { rgb: 'D1FAE5' } }, border: cellBorder };
            const statusBaixo = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: '92400E' } }, fill: { fgColor: { rgb: 'FEF3C7' } }, border: cellBorder };
            const statusZerado = { ...cellCenter, font: { sz: 9, bold: true, color: { rgb: '7F1D1D' } }, fill: { fgColor: { rgb: 'FEE2E2' } }, border: cellBorder };
            const statusOkEven = { ...statusOk, fill: { fgColor: { rgb: 'A7F3D0' } } };
            const statusBaixoEven = { ...statusBaixo, fill: { fgColor: { rgb: 'FCD34D' } } };
            const statusZeradoEven = { ...statusZerado, fill: { fgColor: { rgb: 'FECACA' } } };

            const qtyGreen = { ...cellCenter, font: { sz: 10, bold: true, color: { rgb: '059669' } }, border: cellBorder };
            const qtyRed = { ...cellCenter, font: { sz: 10, bold: true, color: { rgb: 'DC2626' } }, border: cellBorder };
            const qtyGreenEven = { ...qtyGreen, fill: { fgColor: { rgb: 'F0FDF4' } } };
            const qtyRedEven = { ...qtyRed, fill: { fgColor: { rgb: 'FEF2F2' } } };

            // Estilos para caixa de KPI
            const kpiTitleStyle = { font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1F2937' } }, alignment: { horizontal: 'left', vertical: 'center' }, border: cellBorder };
            const kpiValueStyle = { font: { bold: true, sz: 14, color: { rgb: '0066FF' } }, fill: { fgColor: { rgb: 'F0F9FF' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: cellBorder };
            const kpiLabelStyle = { font: { sz: 9, color: { rgb: '6B7280' } }, alignment: { horizontal: 'center', vertical: 'center' }, border: cellBorder, fill: { fgColor: { rgb: 'FFFFFF' } } };

            const applyStyles = (ws, startRow, dataLen, colCount, styleFn) => {
                for (let r = 0; r < dataLen; r++) {
                    for (let c = 0; c < colCount; c++) {
                        const addr = XLS.utils.encode_cell({ r: startRow + r, c });
                        if (ws[addr]) ws[addr].s = styleFn(r, c, ws[addr].v);
                    }
                }
            };

            // ─────────────────────────────────────────────
            // ABA 0: RESUMO EXECUTIVO (Dashboard)
            // ─────────────────────────────────────────────
            const totalProducts = products.length;
            const totalUnits = products.reduce((s, p) => s + (p.quantity || 0), 0);
            const lowCount = products.filter(p => (p.quantity || 0) > 0 && (p.quantity || 0) <= (p.minQuantity || 0)).length;
            const zeroCount = products.filter(p => (p.quantity || 0) === 0).length;
            const okCount = totalProducts - lowCount - zeroCount;

            // Cálculos de movimentação
            const totalEntries = history.filter(h => ['Entrada', 'Entrada por NF', 'Ajuste Entrada', 'Criação'].includes(h.type)).length;
            const totalExits = history.filter(h => ['Saída', 'Saída por Requisição', 'Ajuste Saída'].includes(h.type)).length;
            const lastMonthEntries = history.filter(h => {
                const hDate = h.date ? new Date(h.date.seconds * 1000) : new Date();
                const now = new Date();
                const diff = (now - hDate) / (1000 * 60 * 60 * 24);
                return diff <= 30 && ['Entrada', 'Entrada por NF', 'Ajuste Entrada'].includes(h.type);
            }).length;
            const lastMonthExits = history.filter(h => {
                const hDate = h.date ? new Date(h.date.seconds * 1000) : new Date();
                const now = new Date();
                const diff = (now - hDate) / (1000 * 60 * 60 * 24);
                return diff <= 30 && ['Saída', 'Saída por Requisição', 'Ajuste Saída'].includes(h.type);
            }).length;

            const wsDash = XLS.utils.aoa_to_sheet([
                ['CONTROLE DE ESTOQUE UHE ESTRELA'],
                ['RESUMO EXECUTIVO — DASHBOARD DE GESTÃO'],
                [],
                ['Relatório Gerado em ' + today + ' às ' + hora],
                [],
                ['INDICADORES PRINCIPAIS'],
                [],
                ['Métrica', 'Valor', 'Status'],
                ['Total de Produtos', totalProducts, 'Produtos cadastrados'],
                ['Unidades em Estoque', totalUnits, 'Unidades totais'],
                ['Produtos OK', okCount, 'Acima do mínimo'],
                ['Estoque Baixo', lowCount, 'Abaixo do mínimo'],
                ['Produtos Zerados', zeroCount, 'Sem disponibilidade'],
                [],
                ['MOVIMENTAÇÃO (últimos 30 dias)'],
                [],
                ['Tipo de Movimento', 'Últimos 30 dias', 'Total Histórico'],
                ['Entradas', lastMonthEntries, totalEntries],
                ['Saídas', lastMonthExits, totalExits],
                [],
                ['TAXA DE COBERTURA DE ESTOQUE'],
                [],
                ['Indicador', 'Percentual'],
                ['Produtos com estoque OK', okCount > 0 ? ((okCount / totalProducts) * 100).toFixed(1) + '%' : '0%'],
                ['Produtos em falta crítica', zeroCount > 0 ? ((zeroCount / totalProducts) * 100).toFixed(1) + '%' : '0%']
            ]);

            wsDash['!cols'] = [
                { wch: 35 }, { wch: 18 }, { wch: 30 }
            ];
            wsDash['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
                { s: { r: 5, c: 0 }, e: { r: 5, c: 2 } },
                { s: { r: 14, c: 0 }, e: { r: 14, c: 2 } },
                { s: { r: 20, c: 0 }, e: { r: 20, c: 2 } }
            ];
            
            // Aplicar estilos ao dashboard
            if (wsDash['A1']) wsDash['A1'].s = titleStyle;
            if (wsDash['A2']) wsDash['A2'].s = subtitleStyle;
            
            // Headers de seções
            ['A6', 'A15', 'A21'].forEach(cell => {
                if (wsDash[cell]) wsDash[cell].s = { ...headerStyle, font: { bold: true, sz: 12, color: { rgb: 'FFFFFF' } } };
            });
            
            // Headers de tabelas
            [8, 17, 22].forEach(row => {
                for (let c = 0; c < 3; c++) {
                    const addr = XLS.utils.encode_cell({ r: row, c });
                    if (wsDash[addr]) wsDash[addr].s = headerPurple;
                }
            });

            // Aplicar cores aos dados do dashboard (r=8..12: 5 linhas de dados; r=13 é linha vazia sem célula)
            for (let r = 8; r <= 12; r++) {
                const a0 = XLS.utils.encode_cell({ r, c: 0 });
                const a1 = XLS.utils.encode_cell({ r, c: 1 });
                const a2 = XLS.utils.encode_cell({ r, c: 2 });
                if (wsDash[a0]) wsDash[a0].s = { ...cellBase, font: { bold: true, sz: 11 } };
                if (wsDash[a1]) wsDash[a1].s = { ...cellCenter, font: { bold: true, sz: 12, color: { rgb: '0066FF' } } };
                if (wsDash[a2]) wsDash[a2].s = cellEvenCenter;
            }

            XLS.utils.book_append_sheet(wb, wsDash, 'Dashboard');

            // ─────────────────────────────────────────────
            // ABA 1: ESTOQUE COMPLETO (dados reais)
            // ─────────────────────────────────────────────
            const estoqueHeaders = ['Nº', 'Código SKU', 'Código RM', 'Produto', 'Grupo', 'Unidade', 'Qtd. Atual', 'Qtd. Mínima', 'Falta', 'Localização', 'Status'];
            const sortedProducts = [...products].sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            const estoqueRows = sortedProducts.map((p, i) => {
                const qty = p.quantity || 0;
                const min = p.minQuantity || 0;
                const falta = Math.max(0, min - qty);
                let status = '✅ OK';
                if (qty === 0) status = '🔴 ZERADO';
                else if (qty <= min) status = '🟡 BAIXO';
                return [i + 1, p.code || '', p.codeRM || '', p.name || '', p.group || '', p.unit || 'Un', qty, min, falta, p.location || '', status];
            });

            const wsE = XLS.utils.aoa_to_sheet([
                ['RELATÓRIO DETALHADO DE ESTOQUE — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total: ${totalProducts} produtos | Unidades: ${totalUnits.toLocaleString('pt-BR')}`],
                [],
                estoqueHeaders,
                ...estoqueRows
            ]);
            wsE['!cols'] = [
                { wch: 5 }, { wch: 14 }, { wch: 14 }, { wch: 36 }, { wch: 20 },
                { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 24 }, { wch: 14 }
            ];
            wsE['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 10 } }
            ];
            wsE['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            wsE['!freeze'] = { xSplit: 0, ySplit: 4 };
            
            if (wsE['A1']) wsE['A1'].s = titleStyle;
            if (wsE['A2']) wsE['A2'].s = subtitleStyle;
            for (let c = 0; c < estoqueHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsE[addr]) wsE[addr].s = headerStyle;
            }
            applyStyles(wsE, 4, estoqueRows.length, estoqueHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 10) {
                    if (String(v).includes('ZERADO')) return isEven ? statusZeradoEven : statusZerado;
                    if (String(v).includes('BAIXO')) return isEven ? statusBaixoEven : statusBaixo;
                    return isEven ? statusOkEven : statusOk;
                }
                if (c === 6 || c === 7 || c === 8) {
                    const val = parseFloat(v) || 0;
                    if (c === 8 && val > 0) return isEven ? { ...qtyRedEven, font: { bold: true, sz: 10, color: { rgb: 'DC2626' } } } : { ...qtyRed, font: { bold: true, sz: 10, color: { rgb: 'DC2626' } } };
                    return isEven ? cellEvenCenter : cellCenter;
                }
                if (c === 0 || c === 1 || c === 2) return isEven ? cellEvenBold : cellBold;
                return isEven ? cellEven : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsE, 'Estoque');

            // ─────────────────────────────────────────────
            // ABA 2: ENTRADAS (histórico completo)
            // ─────────────────────────────────────────────
            const entradas = history
                .filter(h => ['Entrada', 'Entrada por NF', 'Ajuste Entrada', 'Criação', 'Importação'].includes(h.type))
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            const entradaHeaders = ['Nº', 'Data', 'Tipo', 'Produto', 'Código RM', 'Quantidade', 'Novo Saldo', 'NF/Lote', 'Fornecedor', 'Registrado Por'];
            const entradaRows = entradas.map((e, i) => {
                const dateStr = e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '';
                return [
                    i + 1, dateStr, e.type || '', e.productName || '', e.productCodeRM || '',
                    e.quantity || 0, e.newTotal ?? '', e.nfNumber || e.details || '',
                    e.supplier || '', e.performedBy || ''
                ];
            });

            const titleGreenStyle = { ...titleStyle, fill: { fgColor: { rgb: '059669' } }, border: { bottom: { style: 'medium', color: { rgb: '047857' } } } };
            const wsEnt = XLS.utils.aoa_to_sheet([
                ['REGISTRO DE ENTRADAS — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total de registros: ${entradas.length}`],
                [],
                entradaHeaders,
                ...entradaRows
            ]);
            wsEnt['!cols'] = [
                { wch: 5 }, { wch: 12 }, { wch: 18 }, { wch: 36 }, { wch: 14 },
                { wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 28 }, { wch: 20 }
            ];
            wsEnt['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
            ];
            wsEnt['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            wsEnt['!freeze'] = { xSplit: 0, ySplit: 4 };
            
            if (wsEnt['A1']) wsEnt['A1'].s = titleGreenStyle;
            if (wsEnt['A2']) wsEnt['A2'].s = subtitleStyle;
            for (let c = 0; c < entradaHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsEnt[addr]) wsEnt[addr].s = headerGreen;
            }
            applyStyles(wsEnt, 4, entradaRows.length, entradaHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 5) return isEven ? qtyGreenEven : qtyGreen;
                if (c === 0 || c === 1 || c === 6) return isEven ? cellEvenGreenCenter : cellEvenGreenCenter;
                return isEven ? cellEvenGreen : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsEnt, 'Entradas');

            // ─────────────────────────────────────────────
            // ABA 3: SAÍDAS (histórico completo)
            // ─────────────────────────────────────────────
            const saidas = history
                .filter(h => ['Saída', 'Saída por Requisição', 'Ajuste Saída'].includes(h.type))
                .sort((a, b) => (b.date?.seconds || 0) - (a.date?.seconds || 0));

            const saidaHeaders = ['Nº', 'Data', 'Produto', 'Código RM', 'Quantidade', 'Novo Saldo', 'Retirado Por', 'Função do Funcionário', 'Local Aplicação', 'Obra', 'Registrado Por'];
            const saidaRows = saidas.map((e, i) => {
                const dateStr = e.date ? new Date(e.date.seconds * 1000).toLocaleDateString('pt-BR') : '';
                return [
                    i + 1, dateStr, e.productName || '', e.productCodeRM || '',
                    e.quantity || 0, e.newTotal ?? '', e.withdrawnBy || '',
                    e.teamLeader || '', e.applicationLocation || '', e.obra || '', e.performedBy || ''
                ];
            });

            const titleRedStyle = { ...titleStyle, fill: { fgColor: { rgb: 'DC2626' } }, border: { bottom: { style: 'medium', color: { rgb: 'B91C1C' } } } };
            const wsSai = XLS.utils.aoa_to_sheet([
                ['REGISTRO DE SAÍDAS — UHE ESTRELA'],
                [`Gerado em ${today} às ${hora} | Total de registros: ${saidas.length}`],
                [],
                saidaHeaders,
                ...saidaRows
            ]);
            wsSai['!cols'] = [
                { wch: 5 }, { wch: 12 }, { wch: 36 }, { wch: 14 }, { wch: 12 },
                { wch: 12 }, { wch: 22 }, { wch: 18 }, { wch: 28 }, { wch: 15 }, { wch: 20 }
            ];
            wsSai['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 10 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 10 } }
            ];
            wsSai['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }, { hpt: 28 }];
            wsSai['!freeze'] = { xSplit: 0, ySplit: 4 };
            
            if (wsSai['A1']) wsSai['A1'].s = titleRedStyle;
            if (wsSai['A2']) wsSai['A2'].s = subtitleStyle;
            for (let c = 0; c < saidaHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 3, c });
                if (wsSai[addr]) wsSai[addr].s = headerRed;
            }
            applyStyles(wsSai, 4, saidaRows.length, saidaHeaders.length, (r, c, v) => {
                const isEven = r % 2 === 1;
                if (c === 4) return isEven ? qtyRedEven : qtyRed;
                if (c === 0 || c === 1 || c === 5 || c === 9) return isEven ? cellEvenRedCenter : cellEvenRedCenter;
                return isEven ? cellEvenRed : cellBase;
            });
            XLS.utils.book_append_sheet(wb, wsSai, 'Saídas');

            // ─────────────────────────────────────────────
            // ABA 4: ANÁLISE & RELATÓRIO (Summary)
            // ─────────────────────────────────────────────
            const groupStats = {};
            products.forEach(p => {
                const group = p.group || 'Sem Grupo';
                if (!groupStats[group]) groupStats[group] = { total: 0, low: 0, zero: 0, units: 0 };
                groupStats[group].total++;
                groupStats[group].units += p.quantity || 0;
                if ((p.quantity || 0) === 0) groupStats[group].zero++;
                else if ((p.quantity || 0) <= (p.minQuantity || 0)) groupStats[group].low++;
            });

            const groupAnalysisRows = Object.entries(groupStats).map(([group, stats]) => [
                group,
                stats.total,
                stats.zero,
                stats.low,
                stats.total - stats.zero - stats.low,
                stats.units
            ]);

            const analysisHeaders = ['Grupo de Produto', 'Total Produtos', 'Zerados', 'Em Falta', 'OK', 'Unidades'];
            
            const wsAnalysis = XLS.utils.aoa_to_sheet([
                ['ANÁLISE & RELATÓRIO DE GESTÃO'],
                [`Análise consolidada em ${today} às ${hora}`],
                [],
                ['RESUMO POR GRUPO DE PRODUTO'],
                [],
                analysisHeaders,
                ...groupAnalysisRows,
                [],
                ['ALERTAS E RECOMENDAÇÕES'],
                []
            ]);

            // Calcular linha de início dinâmica:
            // r=0 título, r=1 subtítulo, r=2 [], r=3 'RESUMO', r=4 [], r=5 headers,
            // r=6..5+N group rows, r=6+N [], r=7+N 'ALERTAS', r=8+N []
            const N = groupAnalysisRows.length;
            const alertasLabelRow = 7 + N; // linha de 'ALERTAS E RECOMENDAÇÕES'
            const alertRow = 9 + N;        // linha onde os alertas começam (após o [] em 8+N)
            const alerts = [];
            
            if (zeroCount > 0) alerts.push(`⚠️ ${zeroCount} produto(s) com estoque ZERADO - Reposição crítica necessária`);
            if (lowCount > 0) alerts.push(`⚠️ ${lowCount} produto(s) com estoque BAIXO - Verifique mínimos`);
            if (lowCount + zeroCount === 0) alerts.push(`✅ Todos os produtos com estoque dentro dos limites`);
            
            alerts.forEach((alert, idx) => {
                wsAnalysis[XLS.utils.encode_cell({ r: alertRow + idx, c: 0 })] = { v: alert, t: 's' };
            });

            // Expandir !ref para cobrir as células de alerta adicionadas manualmente
            const lastAlertRow = alertRow + Math.max(alerts.length - 1, 0);
            wsAnalysis['!ref'] = XLS.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastAlertRow, c: 5 } });

            wsAnalysis['!cols'] = [
                { wch: 28 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }
            ];
            wsAnalysis['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
                { s: { r: 1, c: 0 }, e: { r: 1, c: 5 } },
                { s: { r: 3, c: 0 }, e: { r: 3, c: 5 } },
                { s: { r: alertasLabelRow, c: 0 }, e: { r: alertasLabelRow, c: 5 } }
            ];
            wsAnalysis['!rows'] = [{ hpt: 32 }, { hpt: 20 }, { hpt: 8 }];
            
            if (wsAnalysis['A1']) wsAnalysis['A1'].s = { ...titleStyle, fill: { fgColor: { rgb: '7C3AED' } }, border: { bottom: { style: 'medium', color: { rgb: '6D28D9' } } } };
            if (wsAnalysis['A2']) wsAnalysis['A2'].s = subtitleStyle;
            if (wsAnalysis['A4']) wsAnalysis['A4'].s = { ...headerPurple, font: { bold: true, sz: 12 } };
            // Estilizar 'ALERTAS E RECOMENDAÇÕES' na linha dinâmica correta
            const alertasLabelAddr = XLS.utils.encode_cell({ r: alertasLabelRow, c: 0 });
            if (wsAnalysis[alertasLabelAddr]) wsAnalysis[alertasLabelAddr].s = { ...headerPurple, font: { bold: true, sz: 12 } };
            
            for (let c = 0; c < analysisHeaders.length; c++) {
                const addr = XLS.utils.encode_cell({ r: 5, c });
                if (wsAnalysis[addr]) wsAnalysis[addr].s = headerPurple;
            }

            for (let r = 0; r < groupAnalysisRows.length; r++) {
                const isEven = r % 2 === 1;
                for (let c = 0; c < 6; c++) {
                    const addr = XLS.utils.encode_cell({ r: r + 6, c });
                    if (wsAnalysis[addr]) {
                        if (c === 1 || c === 5) {
                            wsAnalysis[addr].s = isEven ? cellEvenPurpleCenter : cellEvenCenter;
                        } else {
                            wsAnalysis[addr].s = isEven ? cellEvenPurple : cellBase;
                        }
                    }
                }
            }

            XLS.utils.book_append_sheet(wb, wsAnalysis, 'Análise');

            // ─── GERAR E BAIXAR ───
            const stamp = new Date().toISOString().replace(/[:T]/g, '-').slice(0, 16);
            const filename = `Controle_Estoque_UHE_Estrela_${stamp}.xlsx`;
            XLS.writeFile(wb, filename);
            showToast(`✅ Planilha exportada com sucesso: ${filename}`);
            } catch (err) {
                console.error('Erro ao gerar planilha Excel:', err);
                showToast(`Erro ao gerar planilha: ${err.message}`, true);
            }
        };
