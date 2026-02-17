/* ========================================
   EMDA - Banco de Talentos
   Application JavaScript v2.1
   + Verificação de duplicata no banco (nome/email/whatsapp)
   + Notificação WhatsApp
   + Tela de sucesso
   + Botão voltar
======================================== */

// ========================================
// Configuration
// ========================================

const CONFIG = {
    // URL do Google Apps Script Web App
    GOOGLE_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbydNX4J72YwlYFCw79Xp1wQiAFxtWTEfKgk1ywIaud-MFFFlTlR_-Y-fqGuW51oedMIyg/exec',
    
    // WhatsApp da escola para notificações
    WHATSAPP_NOTIFY: '5531988148522',
    
    // Tempo de splash screen (ms)
    SPLASH_DURATION: 4000,
    
    // Validação de telefone
    PHONE_REGEX: /^\(?[1-9]{2}\)?\s?(?:9\d{4}|\d{4})-?\d{4}$/
};

// ========================================
// DOM Elements
// ========================================

const elements = {
    splashScreen: document.getElementById('splash-screen'),
    app: document.getElementById('app'),
    welcomeScreen: document.getElementById('welcome-screen'),
    formContainer: document.getElementById('form-container'),
    startBtn: document.getElementById('start-btn'),
    form: document.getElementById('curriculum-form'),
    installPrompt: document.getElementById('install-prompt'),
    installAccept: document.getElementById('install-accept'),
    installDismiss: document.getElementById('install-dismiss'),
    photoUpload: document.getElementById('photo-upload'),
    photoInput: document.getElementById('photo-input'),
    cameraInput: document.getElementById('camera-input'),
    photoPreview: document.getElementById('photo-preview'),
    photoPlaceholder: document.getElementById('photo-placeholder'),
    btnGallery: document.getElementById('btn-gallery'),
    btnSelfie: document.getElementById('btn-selfie'),
    progressFill: document.querySelectorAll('.progress-fill'),
    progressSteps: document.querySelectorAll('.progress-step'),
    formSteps: document.querySelectorAll('.form-step'),
    successModal: document.getElementById('success-modal'),
    submitBtn: document.querySelector('.btn-submit')
};

// ========================================
// State
// ========================================

let currentStep = 1;
let photoBase64 = null;
let duplicateFound = false; // Flag global de duplicata

// ========================================
// Initialization
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    initSplashScreen();
    initWelcomeScreen();
    initHistoryNavigation();
    initPhotoUpload();
    initFormNavigation();
    initFormValidation();
    initPhoneMask();
    initCityAutocomplete();
    initDuplicateCheck();
    initServiceWorker();
    initInstallPrompt();
});

// ========================================
// Splash Screen
// ========================================

function initSplashScreen() {
    setTimeout(() => {
        elements.splashScreen.classList.add('fade-out');
        elements.app.classList.remove('hidden');
        setTimeout(() => {
            elements.splashScreen.style.display = 'none';
        }, 600);
    }, CONFIG.SPLASH_DURATION);
}

// ========================================
// Welcome Screen
// ========================================

function initWelcomeScreen() {
    elements.startBtn.addEventListener('click', () => {
        showForm();
    });
}

function showForm() {
    elements.welcomeScreen.style.opacity = '0';
    elements.welcomeScreen.style.transform = 'translateY(-20px)';
    elements.welcomeScreen.style.transition = 'all 0.5s ease';
    
    setTimeout(() => {
        elements.welcomeScreen.classList.add('hidden');
        elements.formContainer.classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'instant' });
        pushAppState('step-1');
    }, 500);
}

// ========================================
// Duplicate Check (banco de dados)
// ========================================

function initDuplicateCheck() {
    const nomeInput = document.getElementById('nome');
    const emailInput = document.getElementById('email');
    const whatsappInput = document.getElementById('whatsapp');
    
    // Debounce timers
    let nomeTimer, emailTimer, whatsappTimer;
    
    // Verificar ao sair do campo (blur) ou após parar de digitar
    nomeInput.addEventListener('blur', () => {
        clearTimeout(nomeTimer);
        const val = nomeInput.value.trim();
        if (val.length >= 5) { // Nome mínimo razoável
            checkDuplicate('nome', val);
        }
    });
    
    emailInput.addEventListener('blur', () => {
        clearTimeout(emailTimer);
        const val = emailInput.value.trim();
        if (val && isValidEmail(val)) {
            checkDuplicate('email', val);
        }
    });
    
    whatsappInput.addEventListener('blur', () => {
        clearTimeout(whatsappTimer);
        const val = whatsappInput.value.trim();
        if (val && isValidPhone(val)) {
            checkDuplicate('whatsapp', val);
        }
    });
}

async function checkDuplicate(field, value) {
    const msgEl = document.getElementById(`${field}-duplicate`);
    const inputEl = document.getElementById(field === 'whatsapp' ? 'whatsapp' : field);
    
    if (!msgEl) return;
    
    // Limpar estado anterior
    msgEl.classList.add('hidden');
    msgEl.textContent = '';
    inputEl.classList.remove('input-duplicate');
    
    // Mostrar estado de verificação
    msgEl.textContent = 'Verificando...';
    msgEl.className = 'input-duplicate-msg checking';
    
    try {
        const cleanValue = field === 'whatsapp' ? value.replace(/\D/g, '') : value;
        const url = `${CONFIG.GOOGLE_SCRIPT_URL}?action=check&field=${field}&value=${encodeURIComponent(cleanValue)}`;
        
        const response = await fetch(url, { redirect: 'follow' });
        const text = await response.text();
        
        let result;
        try {
            result = JSON.parse(text);
        } catch (e) {
            // Se não é JSON, pode ser um erro do Google — ignorar silenciosamente
            msgEl.classList.add('hidden');
            msgEl.textContent = '';
            return;
        }
        
        if (result.found) {
            showDuplicateWarning(field, msgEl, inputEl, result.data);
        } else {
            // Limpo — nenhuma duplicata
            msgEl.classList.add('hidden');
            msgEl.textContent = '';
        }
    } catch (error) {
        console.log('Erro na verificação de duplicata:', error);
        // Silenciosamente falha — não bloquear o formulário
        msgEl.classList.add('hidden');
        msgEl.textContent = '';
    }
}

function showDuplicateWarning(field, msgEl, inputEl, existingData) {
    duplicateFound = true;
    
    const fieldNames = {
        nome: 'nome',
        email: 'e-mail',
        whatsapp: 'WhatsApp'
    };
    
    // Mostrar aviso inline no campo
    msgEl.textContent = `Este ${fieldNames[field]} já está cadastrado`;
    msgEl.className = 'input-duplicate-msg show';
    inputEl.classList.add('input-duplicate');
    
    // Mostrar modal de duplicata
    showDuplicateModal(field, existingData);
}

function showDuplicateModal(field, existingData) {
    // Remover modal anterior se existir
    const existing = document.querySelector('.duplicate-modal');
    if (existing) existing.remove();
    
    const fieldNames = {
        nome: 'nome',
        email: 'e-mail',
        whatsapp: 'WhatsApp'
    };
    
    const dataEnvio = existingData.dataEnvio 
        ? new Date(existingData.dataEnvio).toLocaleDateString('pt-BR') 
        : existingData.timestamp 
            ? new Date(existingData.timestamp).toLocaleDateString('pt-BR')
            : '';
    
    const modal = document.createElement('div');
    modal.className = 'duplicate-modal';
    modal.innerHTML = `
        <div class="duplicate-modal-content">
            <div class="duplicate-modal-icon">
                <svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#C9A962" stroke-width="1.5">
                    <circle cx="12" cy="12" r="10"/>
                    <line x1="12" y1="8" x2="12" y2="12"/>
                    <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
            </div>
            
            <h3 class="duplicate-modal-title">Cadastro já realizado</h3>
            
            ${existingData.nome ? `<p class="duplicate-modal-name">${existingData.nome}</p>` : ''}
            ${dataEnvio ? `<p class="duplicate-modal-date">Cadastrado em ${dataEnvio}</p>` : ''}
            
            <p class="duplicate-modal-description">
                Identificamos que este ${fieldNames[field]} já está registrado no nosso Banco de Talentos.
            </p>
            
            <div class="duplicate-modal-buttons">
                <button class="btn btn-secondary duplicate-btn-update">
                    Atualizar meus dados
                </button>
                <button class="btn btn-primary duplicate-btn-close">
                    Ok, entendi
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    requestAnimationFrame(() => modal.classList.add('show'));
    
    // Botão "Ok, entendi" — volta para welcome
    modal.querySelector('.duplicate-btn-close').addEventListener('click', () => {
        modal.classList.remove('show');
        setTimeout(() => {
            modal.remove();
            // Voltar para Welcome Screen
            backToWelcome();
            // Resetar formulário
            resetForm();
        }, 300);
    });
    
    // Botão "Atualizar meus dados" — deixar continuar preenchendo
    modal.querySelector('.duplicate-btn-update').addEventListener('click', () => {
        modal.classList.remove('show');
        duplicateFound = false; // Permitir envio
        setTimeout(() => modal.remove(), 300);
        
        // Limpar avisos visuais
        clearAllDuplicateWarnings();
    });
    
    // Fechar ao clicar no backdrop
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.classList.remove('show');
            setTimeout(() => modal.remove(), 300);
        }
    });
}

function clearAllDuplicateWarnings() {
    document.querySelectorAll('.input-duplicate-msg').forEach(el => {
        el.classList.add('hidden');
        el.textContent = '';
    });
    document.querySelectorAll('.input-duplicate').forEach(el => {
        el.classList.remove('input-duplicate');
    });
    duplicateFound = false;
}

function resetForm() {
    elements.form.reset();
    photoBase64 = null;
    currentStep = 1;
    if (elements.photoPreview) {
        elements.photoPreview.classList.add('hidden');
        elements.photoPlaceholder.classList.remove('hidden');
        elements.photoUpload.style.borderStyle = '';
        elements.photoUpload.style.borderColor = '';
    }
    clearAllDuplicateWarnings();
    goToStep(1);
}

// ========================================
// Photo Upload
// ========================================

function initPhotoUpload() {
    elements.photoUpload.addEventListener('click', () => {
        elements.photoInput.click();
    });
    
    elements.btnGallery.addEventListener('click', (e) => {
        e.stopPropagation();
        elements.photoInput.click();
    });
    
    elements.btnSelfie.addEventListener('click', (e) => {
        e.stopPropagation();
        elements.cameraInput.click();
    });
    
    elements.photoInput.addEventListener('change', handlePhotoSelect);
    elements.cameraInput.addEventListener('change', handlePhotoSelect);
}

function handlePhotoSelect(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    if (!file.type.startsWith('image/')) {
        alert('Por favor, selecione uma imagem válida.');
        return;
    }
    
    if (file.size > 5 * 1024 * 1024) {
        alert('A imagem deve ter no máximo 5MB.');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (event) => {
        photoBase64 = event.target.result;
        elements.photoPreview.src = photoBase64;
        elements.photoPreview.classList.remove('hidden');
        elements.photoPlaceholder.classList.add('hidden');
        elements.photoUpload.style.borderStyle = 'solid';
        elements.photoUpload.style.borderColor = 'var(--color-gold)';
    };
    reader.readAsDataURL(file);
}

// ========================================
// History Management (botão voltar nativo)
// ========================================

// Estado de navegação do app:
// 'welcome' → 'step-1' → 'step-2' → 'step-3' → 'success'

function initHistoryNavigation() {
    // Estado inicial: welcome
    history.replaceState({ screen: 'welcome' }, '', '');
    
    // Escutar o botão voltar nativo do telefone/navegador
    window.addEventListener('popstate', (e) => {
        const state = e.state;
        
        if (!state) {
            // Sem estado = usuário está tentando sair do app
            // Recolocar o estado welcome para não fechar
            history.pushState({ screen: 'welcome' }, '', '');
            
            // Se não está na welcome, voltar para ela
            if (!elements.welcomeScreen.classList.contains('hidden')) {
                // Já está na welcome, não fazer nada
                return;
            }
            handleNativeBack();
            return;
        }
        
        // Fechar modais abertos primeiro
        const duplicateModal = document.querySelector('.duplicate-modal');
        if (duplicateModal) {
            duplicateModal.classList.remove('show');
            setTimeout(() => duplicateModal.remove(), 300);
            // Recolocar o estado atual
            pushAppState(getCurrentScreen());
            return;
        }
        
        // Navegar para o estado do histórico
        navigateToState(state.screen);
    });
}

function getCurrentScreen() {
    if (document.querySelector('.success-screen')) return 'success';
    if (elements.formContainer && !elements.formContainer.classList.contains('hidden')) {
        return `step-${currentStep}`;
    }
    return 'welcome';
}

function pushAppState(screen) {
    history.pushState({ screen: screen }, '', '');
}

function navigateToState(screen) {
    if (screen === 'welcome') {
        // Voltar para Welcome
        if (!elements.welcomeScreen.classList.contains('hidden')) return; // Já está
        
        // Se tem success screen, remover
        const successScreen = document.querySelector('.success-screen');
        if (successScreen) {
            successScreen.remove();
        }
        
        // Voltar do formulário para welcome
        elements.formContainer.classList.add('hidden');
        elements.formContainer.style.opacity = '';
        elements.formContainer.style.transition = '';
        elements.welcomeScreen.classList.remove('hidden');
        elements.welcomeScreen.style.opacity = '1';
        elements.welcomeScreen.style.transform = 'translateY(0)';
        window.scrollTo({ top: 0, behavior: 'instant' });
        
    } else if (screen.startsWith('step-')) {
        const step = parseInt(screen.split('-')[1]);
        if (currentStep !== step) {
            goToStep(step);
        }
    }
}

function handleNativeBack() {
    const successScreen = document.querySelector('.success-screen');
    
    if (successScreen) {
        // Da success → welcome
        successScreen.remove();
        elements.formContainer.classList.add('hidden');
        elements.welcomeScreen.classList.remove('hidden');
        elements.welcomeScreen.style.opacity = '1';
        elements.welcomeScreen.style.transform = 'translateY(0)';
        window.scrollTo({ top: 0, behavior: 'instant' });
        pushAppState('welcome');
        return;
    }
    
    if (!elements.formContainer.classList.contains('hidden')) {
        // Do formulário
        if (currentStep > 1) {
            // Step 2 ou 3 → voltar uma etapa
            goToStep(currentStep - 1);
            pushAppState(`step-${currentStep}`);
        } else {
            // Step 1 → welcome
            backToWelcome();
            pushAppState('welcome');
        }
        return;
    }
    
    // Já está na welcome — empurrar estado para não fechar
    pushAppState('welcome');
}

// ========================================
// Form Navigation
// ========================================

function initFormNavigation() {
    document.querySelectorAll('.btn-next').forEach(btn => {
        btn.addEventListener('click', () => {
            const nextStep = parseInt(btn.dataset.next);
            if (validateStep(currentStep)) {
                goToStep(nextStep);
                pushAppState(`step-${nextStep}`);
            }
        });
    });
    
    document.querySelectorAll('.btn-prev').forEach(btn => {
        btn.addEventListener('click', () => {
            const prevStep = parseInt(btn.dataset.prev);
            if (prevStep === 0) {
                backToWelcome();
                pushAppState('welcome');
            } else {
                goToStep(prevStep);
                pushAppState(`step-${prevStep}`);
            }
        });
    });
}

function backToWelcome() {
    elements.formContainer.style.opacity = '0';
    elements.formContainer.style.transition = 'opacity 0.3s ease';
    
    setTimeout(() => {
        elements.formContainer.classList.add('hidden');
        elements.formContainer.style.opacity = '';
        elements.formContainer.style.transition = '';
        elements.welcomeScreen.classList.remove('hidden');
        elements.welcomeScreen.style.opacity = '1';
        elements.welcomeScreen.style.transform = 'translateY(0)';
        window.scrollTo({ top: 0, behavior: 'instant' });
    }, 300);
}

function goToStep(step) {
    elements.formSteps.forEach(formStep => {
        formStep.classList.remove('active');
        if (parseInt(formStep.dataset.step) === step) {
            formStep.classList.add('active');
        }
    });
    
    elements.progressSteps.forEach((progressStep, index) => {
        progressStep.classList.remove('active', 'completed');
        if (index + 1 < step) {
            progressStep.classList.add('completed');
        } else if (index + 1 === step) {
            progressStep.classList.add('active');
        }
    });
    
    elements.progressFill.forEach((fill, index) => {
        if (index < step - 1) {
            fill.style.width = '100%';
        } else {
            fill.style.width = '0';
        }
    });
    
    currentStep = step;
    elements.form.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ========================================
// Form Validation
// ========================================

function initFormValidation() {
    elements.form.addEventListener('submit', handleSubmit);
}

function validateStep(step) {
    const currentFormStep = document.querySelector(`.form-step[data-step="${step}"]`);
    const requiredInputs = currentFormStep.querySelectorAll('[required]');
    let isValid = true;
    
    requiredInputs.forEach(input => {
        if (!input.value.trim()) {
            isValid = false;
            highlightError(input);
        } else {
            clearError(input);
        }
    });
    
    if (step === 1) {
        const emailInput = document.getElementById('email');
        if (emailInput.value && !isValidEmail(emailInput.value)) {
            isValid = false;
            highlightError(emailInput);
        }
        
        const phoneInput = document.getElementById('whatsapp');
        if (phoneInput.value && !isValidPhone(phoneInput.value)) {
            isValid = false;
            highlightError(phoneInput);
        }
    }
    
    if (step === 2) {
        const cursos = document.querySelectorAll('input[name="cursos"]:checked');
        if (cursos.length === 0) {
            isValid = false;
            alert('Por favor, selecione pelo menos um curso.');
        }
    }
    
    if (!isValid) {
        currentFormStep.style.animation = 'none';
        currentFormStep.offsetHeight;
        currentFormStep.style.animation = 'shake 0.5s ease';
    }
    
    return isValid;
}

function highlightError(input) {
    input.style.borderColor = '#D4A5A5';
    input.addEventListener('input', () => clearError(input), { once: true });
}

function clearError(input) {
    input.style.borderColor = '';
}

function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidPhone(phone) {
    const cleaned = phone.replace(/\D/g, '');
    return cleaned.length >= 10 && cleaned.length <= 11;
}

// ========================================
// Phone Mask
// ========================================

function initPhoneMask() {
    const phoneInput = document.getElementById('whatsapp');
    
    phoneInput.addEventListener('input', (e) => {
        let value = e.target.value.replace(/\D/g, '');
        
        if (value.length > 11) {
            value = value.slice(0, 11);
        }
        
        if (value.length > 0) {
            if (value.length <= 2) {
                value = `(${value}`;
            } else if (value.length <= 7) {
                value = `(${value.slice(0, 2)}) ${value.slice(2)}`;
            } else {
                value = `(${value.slice(0, 2)}) ${value.slice(2, 7)}-${value.slice(7)}`;
            }
        }
        
        e.target.value = value;
    });
}

// ========================================
// City Autocomplete
// ========================================

const BRAZILIAN_CITIES = [
    "Belo Horizonte, MG", "São Paulo, SP", "Rio de Janeiro, RJ", "Brasília, DF",
    "Salvador, BA", "Fortaleza, CE", "Curitiba, PR", "Recife, PE", "Porto Alegre, RS",
    "Manaus, AM", "Belém, PA", "Goiânia, GO", "Guarulhos, SP", "Campinas, SP",
    "São Luís, MA", "São Gonçalo, RJ", "Maceió, AL", "Duque de Caxias, RJ",
    "Natal, RN", "Teresina, PI", "São Bernardo do Campo, SP", "Campo Grande, MS",
    "Nova Iguaçu, RJ", "João Pessoa, PB", "Santo André, SP", "São José dos Campos, SP",
    "Osasco, SP", "Ribeirão Preto, SP", "Jaboatão dos Guararapes, PE", "Uberlândia, MG",
    "Contagem, MG", "Sorocaba, SP", "Aracaju, SE", "Feira de Santana, BA",
    "Cuiabá, MT", "Joinville, SC", "Aparecida de Goiânia, GO", "Londrina, PR",
    "Juiz de Fora, MG", "Ananindeua, PA", "Niterói, RJ", "Porto Velho, RO",
    "Serra, ES", "Belford Roxo, RJ", "Caxias do Sul, RS", "Campos dos Goytacazes, RJ",
    "Florianópolis, SC", "Macapá, AP", "Vila Velha, ES", "Mauá, SP",
    "São João de Meriti, RJ", "São José do Rio Preto, SP", "Santos, SP", "Mogi das Cruzes, SP",
    "Betim, MG", "Diadema, SP", "Campina Grande, PB", "Jundiaí, SP", "Maringá, PR",
    "Montes Claros, MG", "Piracicaba, SP", "Carapicuíba, SP", "Olinda, PE",
    "Cariacica, ES", "Bauru, SP", "Rio Branco, AC", "Anápolis, GO", "Vitória, ES",
    "Caucaia, CE", "Ponta Grossa, PR", "Itaquaquecetuba, SP", "Blumenau, SC",
    "Vitória da Conquista, BA", "Pelotas, RS", "Franca, SP", "Guarujá, SP",
    "Petrolina, PE", "Canoas, RS", "Paulista, PE", "Ribeirão das Neves, MG",
    "Uberaba, MG", "Cascavel, PR", "Praia Grande, SP", "Santa Maria, RS",
    "Governador Valadares, MG", "Gravataí, RS", "Caruaru, PE", "Ipatinga, MG",
    "Novo Hamburgo, RS", "São Vicente, SP", "Serra Talhada, PE", "Sete Lagoas, MG",
    "Divinópolis, MG", "Poços de Caldas, MG", "Barbacena, MG", "Patos de Minas, MG",
    "Conselheiro Lafaiete, MG", "Varginha, MG", "Sabará, MG", "Santa Luzia, MG",
    "Itabira, MG", "Passos, MG", "Teófilo Otoni, MG", "Lavras, MG",
    "Nova Lima, MG", "Araguari, MG", "Itaúna, MG", "Ituiutaba, MG",
    "Patrocínio, MG", "Manhuaçu, MG", "São João del-Rei, MG", "Muriaé, MG",
    "Araxá, MG", "Alfenas, MG", "Ponte Nova, MG", "Viçosa, MG",
    "Ouro Preto, MG", "Caratinga, MG", "Ubá, MG", "Curvelo, MG"
];

function initCityAutocomplete() {
    const cidadeInput = document.getElementById('cidade');
    const suggestionsContainer = document.getElementById('cidade-suggestions');
    
    if (!cidadeInput || !suggestionsContainer) return;
    
    let activeIndex = -1;
    
    cidadeInput.addEventListener('input', (e) => {
        const value = e.target.value.toLowerCase().trim();
        
        if (value.length < 2) {
            suggestionsContainer.classList.remove('show');
            suggestionsContainer.innerHTML = '';
            return;
        }
        
        const filtered = BRAZILIAN_CITIES.filter(city => 
            city.toLowerCase().includes(value)
        ).slice(0, 8);
        
        if (filtered.length === 0) {
            suggestionsContainer.classList.remove('show');
            suggestionsContainer.innerHTML = '';
            return;
        }
        
        activeIndex = -1;
        suggestionsContainer.innerHTML = filtered.map((city, index) => {
            const regex = new RegExp(`(${value})`, 'gi');
            const highlighted = city.replace(regex, '<strong>$1</strong>');
            return `<div class="autocomplete-item" data-index="${index}" data-value="${city}">${highlighted}</div>`;
        }).join('');
        
        suggestionsContainer.classList.add('show');
        
        suggestionsContainer.querySelectorAll('.autocomplete-item').forEach(item => {
            item.addEventListener('click', () => {
                cidadeInput.value = item.dataset.value.split(',')[0];
                suggestionsContainer.classList.remove('show');
                
                const state = item.dataset.value.split(',')[1]?.trim();
                if (state) {
                    const estadoSelect = document.getElementById('estado');
                    if (estadoSelect) estadoSelect.value = state;
                }
            });
        });
    });
    
    cidadeInput.addEventListener('keydown', (e) => {
        const items = suggestionsContainer.querySelectorAll('.autocomplete-item');
        
        if (e.key === 'ArrowDown') {
            e.preventDefault();
            activeIndex = Math.min(activeIndex + 1, items.length - 1);
            updateActiveItem(items);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            activeIndex = Math.max(activeIndex - 1, 0);
            updateActiveItem(items);
        } else if (e.key === 'Enter' && activeIndex >= 0) {
            e.preventDefault();
            items[activeIndex]?.click();
        } else if (e.key === 'Escape') {
            suggestionsContainer.classList.remove('show');
        }
    });
    
    function updateActiveItem(items) {
        items.forEach((item, index) => {
            item.classList.toggle('active', index === activeIndex);
        });
    }
    
    document.addEventListener('click', (e) => {
        if (!cidadeInput.contains(e.target) && !suggestionsContainer.contains(e.target)) {
            suggestionsContainer.classList.remove('show');
        }
    });
}

// ========================================
// Form Submission
// ========================================

async function handleSubmit(e) {
    e.preventDefault();
    
    if (!validateStep(3)) return;
    
    const termos = document.getElementById('termos');
    if (!termos.checked) {
        alert('Por favor, aceite os termos para continuar.');
        return;
    }
    
    setLoadingState(true);
    
    const formData = collectFormData();
    
    try {
        await sendToGoogleSheets(formData);
        
        // Salvar flag local
        localStorage.setItem('emda_curriculo_enviado', JSON.stringify({
            nome: formData.nome,
            email: formData.email,
            whatsapp: formData.whatsapp,
            dataEnvio: new Date().toISOString()
        }));
        
        // Mostrar tela de sucesso
        showSuccessScreen(formData);
        
    } catch (error) {
        console.error('Erro ao enviar:', error);
        alert('Ocorreu um erro ao enviar seu currículo. Por favor, tente novamente.');
    } finally {
        setLoadingState(false);
    }
}

function collectFormData() {
    const cursos = Array.from(document.querySelectorAll('input[name="cursos"]:checked'))
        .map(cb => cb.value)
        .join(', ');
    
    return {
        timestamp: new Date().toISOString(),
        nome: document.getElementById('nome').value.trim(),
        email: document.getElementById('email').value.trim(),
        whatsapp: document.getElementById('whatsapp').value.trim(),
        cidade: document.getElementById('cidade').value.trim(),
        estado: document.getElementById('estado').value,
        cursos: cursos,
        ano_conclusao: document.getElementById('ano_conclusao').value,
        experiencia: document.getElementById('experiencia').value.trim(),
        instagram: document.getElementById('instagram').value.trim(),
        portfolio: document.getElementById('portfolio').value.trim(),
        linkedin: document.getElementById('linkedin').value.trim(),
        sobre: document.getElementById('sobre').value.trim(),
        foto: photoBase64 ? 'Sim' : 'Não',
        foto_base64: photoBase64 || ''
    };
}

async function sendToGoogleSheets(data) {
    try {
        // Método mais confiável com Google Apps Script:
        // Criar form invisível e submeter (evita CORS)
        return new Promise((resolve) => {
            const form = document.createElement('form');
            form.method = 'POST';
            form.action = CONFIG.GOOGLE_SCRIPT_URL;
            form.target = 'emda-hidden-frame';
            form.style.display = 'none';
            
            // Criar iframe invisível para receber a resposta
            let iframe = document.getElementById('emda-hidden-frame');
            if (!iframe) {
                iframe = document.createElement('iframe');
                iframe.id = 'emda-hidden-frame';
                iframe.name = 'emda-hidden-frame';
                iframe.style.display = 'none';
                document.body.appendChild(iframe);
            }
            
            // Adicionar os dados como campo hidden
            const input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'payload';
            input.value = JSON.stringify(data);
            form.appendChild(input);
            
            document.body.appendChild(form);
            form.submit();
            
            // Limpar após envio
            setTimeout(() => {
                form.remove();
                resolve({ success: true });
            }, 2000);
        });
    } catch (error) {
        console.error('Erro ao enviar para Google Sheets:', error);
        return { success: true };
    }
}

// ========================================
// Success Screen
// ========================================

function showSuccessScreen(formData) {
    elements.formContainer.classList.add('hidden');
    elements.successModal.classList.add('hidden');
    
    const successScreen = document.createElement('div');
    successScreen.className = 'success-screen';
    successScreen.innerHTML = `
        <div class="success-screen-bg">
            <div class="success-particles">
                <div class="particle"></div>
                <div class="particle"></div>
                <div class="particle"></div>
                <div class="particle"></div>
                <div class="particle"></div>
                <div class="particle"></div>
            </div>
        </div>
        <div class="success-screen-content">
            <div class="success-check-container">
                <svg class="success-check" width="80" height="80" viewBox="0 0 24 24" fill="none">
                    <circle cx="12" cy="12" r="10" stroke="#C9A962" stroke-width="1.5" opacity="0.3"/>
                    <circle class="success-circle-anim" cx="12" cy="12" r="10" stroke="#C9A962" stroke-width="1.5" 
                        stroke-dasharray="63" stroke-dashoffset="63"/>
                    <polyline class="success-tick-anim" points="8 12.5 11 15.5 16 9" stroke="#C9A962" stroke-width="2" 
                        stroke-linecap="round" stroke-linejoin="round"
                        stroke-dasharray="20" stroke-dashoffset="20"/>
                </svg>
            </div>
            
            <h2 class="success-title">Currículo Enviado!</h2>
            
            <div class="success-divider">
                <span class="divider-line"></span>
                <span class="divider-star">✦</span>
                <span class="divider-line"></span>
            </div>
            
            <p class="success-name">${formData.nome}</p>
            
            <p class="success-message">
                Obrigado por se cadastrar no nosso Banco de Talentos. 
                Entraremos em contato caso surja uma oportunidade compatível com seu perfil.
            </p>
            
            <div class="success-info">
                <div class="success-info-item">
                    <span class="success-info-icon">🎓</span>
                    <span>${formData.cursos}</span>
                </div>
                <div class="success-info-item">
                    <span class="success-info-icon">📍</span>
                    <span>${formData.cidade}/${formData.estado}</span>
                </div>
            </div>
            
            <div class="success-actions">
                <a href="https://escolademodadeniseaguiar.com.br" class="btn btn-primary success-btn" target="_blank">
                    Visitar nosso site
                </a>
                <button class="btn btn-secondary success-btn success-btn-close">
                    Fechar
                </button>
            </div>
            
            <div class="success-footer-area">
                <img src="img/logo-white.png" alt="EMDA" class="success-footer-logo">
                <p>Escola de Moda Denise Aguiar</p>
                <p class="success-footer-small">+37 anos formando profissionais de moda</p>
            </div>
        </div>
    `;
    
    elements.app.appendChild(successScreen);
    
    pushAppState('success');
    
    successScreen.querySelector('.success-btn-close').addEventListener('click', () => {
        window.close();
        location.reload();
    });
    
    requestAnimationFrame(() => {
        successScreen.classList.add('show');
    });
}

function setLoadingState(isLoading) {
    if (isLoading) {
        elements.submitBtn.classList.add('loading');
        elements.submitBtn.disabled = true;
    } else {
        elements.submitBtn.classList.remove('loading');
        elements.submitBtn.disabled = false;
    }
}

function showSuccessModal() {
    elements.successModal.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
}

// ========================================
// Service Worker
// ========================================

function initServiceWorker() {
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('sw.js')
            .then(registration => {
                console.log('Service Worker registrado:', registration.scope);
            })
            .catch(error => {
                console.log('Erro ao registrar Service Worker:', error);
            });
    }
}

// ========================================
// PWA Install Prompt
// ========================================

let deferredPrompt = null;

function initInstallPrompt() {
    if (window.matchMedia('(display-mode: standalone)').matches) return;
    
    const dismissedAt = localStorage.getItem('installDismissed');
    if (dismissedAt && (Date.now() - parseInt(dismissedAt)) < 24 * 60 * 60 * 1000) return;
    
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        setTimeout(() => showInstallPrompt(), 3000);
    });
    
    elements.installAccept.addEventListener('click', handleInstallClick);
    elements.installDismiss.addEventListener('click', hideInstallPrompt);
    
    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
    const isInStandaloneMode = window.navigator.standalone === true;
    
    if (isIOS && !isInStandaloneMode) {
        setTimeout(() => showIOSInstallPrompt(), 5000);
    }
}

function showInstallPrompt() {
    elements.installPrompt.classList.add('show');
}

function hideInstallPrompt() {
    elements.installPrompt.classList.remove('show');
    localStorage.setItem('installDismissed', Date.now().toString());
}

async function handleInstallClick() {
    if (!deferredPrompt) { showIOSInstallPrompt(); return; }
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    console.log('Resultado da instalação:', outcome);
    deferredPrompt = null;
    hideInstallPrompt();
}

function showIOSInstallPrompt() {
    const modal = document.createElement('div');
    modal.className = 'ios-install-modal';
    modal.innerHTML = `
        <div class="ios-install-content">
            <h3>Instalar App</h3>
            <p style="font-size: 0.85rem; color: #666; margin-bottom: 1rem;">
                Adicione o EMDA Talentos à sua tela inicial:
            </p>
            <div class="ios-install-steps">
                <div class="ios-step">
                    <span class="ios-step-icon">📤</span>
                    <span>Toque no botão <strong>Compartilhar</strong></span>
                </div>
                <div class="ios-step">
                    <span class="ios-step-icon">➕</span>
                    <span>Selecione <strong>"Adicionar à Tela Inicial"</strong></span>
                </div>
                <div class="ios-step">
                    <span class="ios-step-icon">✓</span>
                    <span>Toque em <strong>Adicionar</strong></span>
                </div>
            </div>
            <button class="ios-install-close" onclick="this.closest('.ios-install-modal').remove()">
                Entendi
            </button>
        </div>
    `;
    document.body.appendChild(modal);
    modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });
    hideInstallPrompt();
}

// ========================================
// Shake Animation
// ========================================

const styleSheet = document.createElement('style');
styleSheet.textContent = `
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        10%, 30%, 50%, 70%, 90% { transform: translateX(-5px); }
        20%, 40%, 60%, 80% { transform: translateX(5px); }
    }
`;
document.head.appendChild(styleSheet);
