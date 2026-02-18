/* ========================================
   EMDA - Conexão Moda
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
    initCursosAnos();
    initDuplicateCheck();
    initAdmin();
    initBiometric();
    initPasswordToggles();
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
// Cursos com Ano de Conclusão Individual
// ========================================

function initCursosAnos() {
    const checkboxes = document.querySelectorAll('input[name="cursos"]');
    const container = document.getElementById('cursos-anos-container');
    const list = document.getElementById('cursos-anos-list');
    
    if (!checkboxes.length || !container || !list) return;
    
    const anosOptions = `
        <option value="">Ano</option>
        <option value="2026">2026</option>
        <option value="2025">2025</option>
        <option value="2024">2024</option>
        <option value="2023">2023</option>
        <option value="2022">2022</option>
        <option value="2021">2021</option>
        <option value="2020">2020</option>
        <option value="2019">2019</option>
        <option value="2018">2018</option>
        <option value="2017">2017</option>
        <option value="2016">2016</option>
        <option value="2015">2015</option>
        <option value="anterior">Anterior a 2015</option>
    `;
    
    checkboxes.forEach(cb => {
        cb.addEventListener('change', () => {
            updateCursosAnos();
        });
    });
    
    function updateCursosAnos() {
        const checked = document.querySelectorAll('input[name="cursos"]:checked');
        
        if (checked.length === 0) {
            container.style.display = 'none';
            list.innerHTML = '';
            return;
        }
        
        container.style.display = 'block';
        
        // Preservar anos já selecionados
        const existingAnos = {};
        list.querySelectorAll('.curso-ano-item').forEach(item => {
            const nome = item.querySelector('.curso-nome').textContent;
            const ano = item.querySelector('select').value;
            existingAnos[nome] = ano;
        });
        
        list.innerHTML = '';
        
        checked.forEach(cb => {
            const cursoNome = cb.value;
            const item = document.createElement('div');
            item.className = 'curso-ano-item';
            item.innerHTML = `
                <span class="curso-nome">${cursoNome}</span>
                <select class="curso-ano-select">${anosOptions}</select>
            `;
            
            // Restaurar ano se já existia
            if (existingAnos[cursoNome]) {
                const select = item.querySelector('select');
                select.value = existingAnos[cursoNome];
            }
            
            list.appendChild(item);
        });
    }
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
                Identificamos que este ${fieldNames[field]} já está registrado no nosso Conexão Moda.
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
    elements.photoUpload.addEventListener('click', (e) => {
        // Só abre galeria se não tem foto ainda
        if (!photoBase64) {
            elements.photoInput.click();
        }
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

// Estado do crop/drag da foto
let photoDrag = {
    isDragging: false,
    startX: 0,
    startY: 0,
    offsetX: 0,
    offsetY: 0,
    currentX: 0,
    currentY: 0,
    scale: 1
};

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
        
        // Resetar posição
        photoDrag.currentX = 0;
        photoDrag.currentY = 0;
        photoDrag.scale = 1;
        updatePhotoTransform();
        
        // Mostrar controles
        const controls = document.getElementById('photo-controls');
        const hint = document.getElementById('photo-drag-hint');
        if (controls) controls.classList.remove('hidden');
        if (hint) hint.classList.remove('hidden');
        
        // Resetar zoom slider
        const zoomSlider = document.getElementById('photo-zoom');
        if (zoomSlider) zoomSlider.value = 100;
        
        // Inicializar drag e zoom
        initPhotoDragZoom();
    };
    reader.readAsDataURL(file);
}

function initPhotoDragZoom() {
    const preview = elements.photoPreview;
    const container = elements.photoUpload;
    const zoomSlider = document.getElementById('photo-zoom');
    const btnZoomIn = document.getElementById('btn-zoom-in');
    const btnZoomOut = document.getElementById('btn-zoom-out');
    
    // Remover listeners antigos (para não duplicar)
    const newPreview = preview.cloneNode(true);
    preview.parentNode.replaceChild(newPreview, preview);
    elements.photoPreview = newPreview;
    
    // --- DRAG (mouse) ---
    newPreview.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        photoDrag.isDragging = true;
        photoDrag.startX = e.clientX - photoDrag.currentX;
        photoDrag.startY = e.clientY - photoDrag.currentY;
        newPreview.classList.add('dragging');
    });
    
    document.addEventListener('mousemove', (e) => {
        if (!photoDrag.isDragging) return;
        photoDrag.currentX = e.clientX - photoDrag.startX;
        photoDrag.currentY = e.clientY - photoDrag.startY;
        updatePhotoTransform();
    });
    
    document.addEventListener('mouseup', () => {
        photoDrag.isDragging = false;
        newPreview.classList.remove('dragging');
    });
    
    // --- DRAG (touch) ---
    newPreview.addEventListener('touchstart', (e) => {
        if (e.touches.length === 1) {
            e.stopPropagation();
            photoDrag.isDragging = true;
            photoDrag.startX = e.touches[0].clientX - photoDrag.currentX;
            photoDrag.startY = e.touches[0].clientY - photoDrag.currentY;
        }
    }, { passive: true });
    
    document.addEventListener('touchmove', (e) => {
        if (!photoDrag.isDragging) return;
        if (e.touches.length === 1) {
            photoDrag.currentX = e.touches[0].clientX - photoDrag.startX;
            photoDrag.currentY = e.touches[0].clientY - photoDrag.startY;
            updatePhotoTransform();
        }
    }, { passive: true });
    
    document.addEventListener('touchend', () => {
        photoDrag.isDragging = false;
    });
    
    // --- ZOOM (slider) ---
    if (zoomSlider) {
        zoomSlider.addEventListener('input', (e) => {
            photoDrag.scale = parseInt(e.target.value) / 100;
            updatePhotoTransform();
        });
    }
    
    // --- ZOOM (botões) ---
    if (btnZoomIn) {
        btnZoomIn.addEventListener('click', (e) => {
            e.stopPropagation();
            photoDrag.scale = Math.min(3, photoDrag.scale + 0.1);
            if (zoomSlider) zoomSlider.value = Math.round(photoDrag.scale * 100);
            updatePhotoTransform();
        });
    }
    
    if (btnZoomOut) {
        btnZoomOut.addEventListener('click', (e) => {
            e.stopPropagation();
            photoDrag.scale = Math.max(1, photoDrag.scale - 0.1);
            if (zoomSlider) zoomSlider.value = Math.round(photoDrag.scale * 100);
            updatePhotoTransform();
        });
    }
    
    // --- PINCH ZOOM (touch) ---
    let lastPinchDist = 0;
    container.addEventListener('touchstart', (e) => {
        if (e.touches.length === 2) {
            lastPinchDist = Math.hypot(
                e.touches[0].clientX - e.touches[1].clientX,
                e.touches[0].clientY - e.touches[1].clientY
            );
        }
    }, { passive: true });
    
    container.addEventListener('touchmove', (e) => {
        if (e.touches.length === 2) {
            const dist = Math.hypot(
                e.touches[0].clientX - e.touches[1].clientX,
                e.touches[0].clientY - e.touches[1].clientY
            );
            const delta = (dist - lastPinchDist) * 0.01;
            photoDrag.scale = Math.max(1, Math.min(3, photoDrag.scale + delta));
            lastPinchDist = dist;
            if (zoomSlider) zoomSlider.value = Math.round(photoDrag.scale * 100);
            updatePhotoTransform();
        }
    }, { passive: true });
}

function updatePhotoTransform() {
    if (elements.photoPreview) {
        elements.photoPreview.style.transform = 
            `translate(${photoDrag.currentX}px, ${photoDrag.currentY}px) scale(${photoDrag.scale})`;
    }
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
        
        // Fechar overlays admin primeiro (maior prioridade)
        const photoFull = document.getElementById('admin-photo-full');
        if (photoFull && !photoFull.classList.contains('hidden')) {
            photoFull.classList.add('hidden');
            pushAppState(getCurrentScreen());
            return;
        }
        
        const deleteConfirm = document.getElementById('admin-delete-confirm');
        if (deleteConfirm && !deleteConfirm.classList.contains('hidden')) {
            deleteConfirm.classList.add('hidden');
            pushAppState(getCurrentScreen());
            return;
        }
        
        const changePw = document.getElementById('admin-change-pw');
        if (changePw && !changePw.classList.contains('hidden')) {
            changePw.classList.add('hidden');
            pushAppState(getCurrentScreen());
            return;
        }
        
        const adminDetail = document.getElementById('admin-detail');
        if (adminDetail && !adminDetail.classList.contains('hidden')) {
            adminDetail.classList.add('hidden');
            pushAppState('admin-panel');
            return;
        }
        
        const adminPanel = document.getElementById('admin-panel');
        if (adminPanel && !adminPanel.classList.contains('hidden')) {
            adminPanel.classList.add('hidden');
            pushAppState('welcome');
            return;
        }
        
        const adminLogin = document.getElementById('admin-login');
        if (adminLogin && !adminLogin.classList.contains('hidden')) {
            adminLogin.classList.add('hidden');
            pushAppState('welcome');
            return;
        }
        
        if (!state) {
            history.pushState({ screen: 'welcome' }, '', '');
            if (!elements.welcomeScreen.classList.contains('hidden')) {
                return;
            }
            handleNativeBack();
            return;
        }
        
        // Fechar modais abertos
        const duplicateModal = document.querySelector('.duplicate-modal');
        if (duplicateModal) {
            duplicateModal.classList.remove('show');
            setTimeout(() => duplicateModal.remove(), 300);
            pushAppState(getCurrentScreen());
            return;
        }
        
        // Navegar para o estado do histórico
        navigateToState(state.screen);
    });
}

function getCurrentScreen() {
    const adminPanel = document.getElementById('admin-panel');
    if (adminPanel && !adminPanel.classList.contains('hidden')) return 'admin-panel';
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
        } else {
            // Verificar se todos os cursos têm ano selecionado
            const anosSelects = document.querySelectorAll('.curso-ano-select');
            let todosComAno = true;
            anosSelects.forEach(sel => {
                if (!sel.value) todosComAno = false;
            });
            if (!todosComAno) {
                isValid = false;
                alert('Por favor, selecione o ano de conclusão para todos os cursos.');
            }
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
    // Coletar cursos com seus anos
    const cursosAnos = Array.from(document.querySelectorAll('.curso-ano-item')).map(item => {
        const nome = item.querySelector('.curso-nome').textContent;
        const ano = item.querySelector('select').value;
        return `${nome} (${ano})`;
    });
    
    const cursos = cursosAnos.join('\n');
    
    return {
        timestamp: new Date().toISOString(),
        nome: document.getElementById('nome').value.trim(),
        email: document.getElementById('email').value.trim(),
        whatsapp: document.getElementById('whatsapp').value.trim(),
        cidade: document.getElementById('cidade').value.trim(),
        estado: document.getElementById('estado').value,
        cursos: cursos,
        experiencia: document.getElementById('experiencia').value.trim(),
        instagram: document.getElementById('instagram').value.trim().replace(/^@+/, ''),
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
                Obrigado por se cadastrar no nosso Conexão Moda. 
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
    
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        // Não mostra automaticamente — será chamado ao abrir gestão
    });
    
    elements.installAccept.addEventListener('click', handleInstallClick);
    elements.installDismiss.addEventListener('click', hideInstallPrompt);
}

function showInstallPromptInAdmin() {
    // Só mostra se não está instalado e não foi dispensado recentemente
    if (window.matchMedia('(display-mode: standalone)').matches) return;
    
    const dismissedAt = localStorage.getItem('installDismissed');
    if (dismissedAt && (Date.now() - parseInt(dismissedAt)) < 24 * 60 * 60 * 1000) return;
    
    if (deferredPrompt) {
        setTimeout(() => showInstallPrompt(), 1000);
    } else {
        const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
        const isInStandaloneMode = window.navigator.standalone === true;
        if (isIOS && !isInStandaloneMode) {
            setTimeout(() => showIOSInstallPrompt(), 1000);
        }
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

// ========================================
// Admin Panel
// ========================================

let ADMIN_PIN = '2026@Tifannypaes';
let adminData = [];
let adminDeleteRow = null;
let adminDeleteName = '';
let adminAuthMode = 'pin'; // 'pin' ou 'bio'

function initAdmin() {
    const accessBtn = document.getElementById('admin-access-btn');
    const loginOverlay = document.getElementById('admin-login');
    const loginClose = document.getElementById('admin-login-close');
    const loginBtn = document.getElementById('admin-login-btn');
    const passwordInput = document.getElementById('admin-password');
    const loginError = document.getElementById('admin-login-error');
    const backBtn = document.getElementById('admin-back-btn');
    const refreshBtn = document.getElementById('admin-refresh-btn');
    const searchInput = document.getElementById('admin-search-input');
    const detailOverlay = document.getElementById('admin-detail');
    const detailClose = document.getElementById('admin-detail-close');
    const deleteConfirm = document.getElementById('admin-delete-confirm');
    const deleteCancel = document.getElementById('admin-delete-cancel');
    const deleteYes = document.getElementById('admin-delete-yes');
    
    if (!accessBtn) return;
    
    // Abrir login
    accessBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        loginOverlay.classList.remove('hidden');
        passwordInput.value = '';
        loginError.classList.add('hidden');
        pushAppState('admin-login');
        setTimeout(() => passwordInput.focus(), 300);
    });
    
    // Fechar login
    loginClose.addEventListener('click', () => {
        loginOverlay.classList.add('hidden');
    });
    
    loginOverlay.addEventListener('click', (e) => {
        if (e.target === loginOverlay) loginOverlay.classList.add('hidden');
    });
    
    // Login
    loginBtn.addEventListener('click', () => adminLogin());
    passwordInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') adminLogin();
    });
    
    async function adminLogin() {
        const pin = passwordInput.value.trim();
        if (!pin) return;
        
        loginBtn.textContent = 'Verificando...';
        loginBtn.disabled = true;
        
        try {
            // Validar no servidor
            const url = CONFIG.GOOGLE_SCRIPT_URL + '?action=list&pin=' + encodeURIComponent(pin);
            const response = await fetch(url);
            const json = await response.json();
            
            if (json.error) {
                loginError.classList.remove('hidden');
                passwordInput.classList.add('input-error');
                setTimeout(() => passwordInput.classList.remove('input-error'), 500);
            } else {
                // Senha correta — atualizar PIN local
                ADMIN_PIN = pin;
                adminAuthMode = 'pin';
                loginOverlay.classList.add('hidden');
                
                // Já temos os dados, passar direto
                document.getElementById('admin-panel').classList.remove('hidden');
                pushAppState('admin-panel');
                adminData = json.data || [];
                document.getElementById('admin-loading').classList.add('hidden');
                if (adminData.length === 0) {
                    document.getElementById('admin-empty').classList.remove('hidden');
                } else {
                    updateAdminStats();
                    renderAdminList('');
                }
                
                // Oferecer biometria
                if (window.PublicKeyCredential && !localStorage.getItem(BIOMETRIC_STORAGE_KEY)) {
                    setTimeout(() => {
                        if (confirm('Deseja ativar a biometria neste dispositivo para acessar mais rápido?')) {
                            biometricRegister();
                        }
                    }, 500);
                }
            }
        } catch (error) {
            loginError.textContent = 'Erro de conexão';
            loginError.classList.remove('hidden');
        } finally {
            loginBtn.textContent = 'Entrar';
            loginBtn.disabled = false;
        }
    }
    
    // Back
    backBtn.addEventListener('click', () => {
        document.getElementById('admin-panel').classList.add('hidden');
        history.back();
    });
    
    // Refresh
    refreshBtn.addEventListener('click', () => loadAdminData());
    
    // Search
    searchInput.addEventListener('input', () => {
        renderAdminList(searchInput.value.trim().toLowerCase());
    });
    
    // Detail close
    detailClose.addEventListener('click', () => {
        detailOverlay.classList.add('hidden');
    });
    
    detailOverlay.addEventListener('click', (e) => {
        if (e.target === detailOverlay) detailOverlay.classList.add('hidden');
    });
    
    // Delete confirm
    deleteCancel.addEventListener('click', () => {
        deleteConfirm.classList.add('hidden');
    });
    
    deleteYes.addEventListener('click', () => adminDeleteConfirmed());
    
    // Settings / Change Password
    const settingsBtn = document.getElementById('admin-settings-btn');
    const changePwOverlay = document.getElementById('admin-change-pw');
    const pwClose = document.getElementById('admin-pw-close');
    const pwBtn = document.getElementById('admin-pw-btn');
    
    settingsBtn.addEventListener('click', () => {
        changePwOverlay.classList.remove('hidden');
        pushAppState('admin-settings');
        document.getElementById('admin-pw-current').value = '';
        document.getElementById('admin-pw-new').value = '';
        document.getElementById('admin-pw-confirm').value = '';
        document.getElementById('admin-pw-error').classList.add('hidden');
        document.getElementById('admin-pw-success').classList.add('hidden');
        updateBiometricSettings();
    });
    
    pwClose.addEventListener('click', () => {
        changePwOverlay.classList.add('hidden');
    });
    
    changePwOverlay.addEventListener('click', (e) => {
        if (e.target === changePwOverlay) changePwOverlay.classList.add('hidden');
    });
    
    pwBtn.addEventListener('click', () => adminChangePassword());
    
    // Inicializar filtro de data
    initDateFilter();
}

function getAdminAuthParam() {
    if (adminAuthMode === 'bio') {
        return 'biotoken=emda-bio-auth';
    }
    return 'pin=' + encodeURIComponent(ADMIN_PIN);
}

function openAdminPanel() {
    document.getElementById('admin-panel').classList.remove('hidden');
    pushAppState('admin-panel');
    // Reset date filter
    activeDateFilter = { type: 'days', days: 7 };
    updateDateFilterLabel('Últimos 7 dias');
    const filterPanel = document.getElementById('admin-date-filter');
    if (filterPanel) {
        filterPanel.classList.add('hidden');
        filterPanel.querySelectorAll('.date-filter-opt').forEach(b => b.classList.remove('active'));
        const first = filterPanel.querySelector('[data-days="7"]');
        if (first) first.classList.add('active');
    }
    document.getElementById('admin-stat-recentes')?.classList.remove('open');
    loadAdminData();
    showInstallPromptInAdmin();
}

async function loadAdminData() {
    const loading = document.getElementById('admin-loading');
    const empty = document.getElementById('admin-empty');
    const list = document.getElementById('admin-list');
    
    loading.classList.remove('hidden');
    empty.classList.add('hidden');
    list.innerHTML = '';
    
    try {
        const url = CONFIG.GOOGLE_SCRIPT_URL + '?action=list&' + getAdminAuthParam();
        const response = await fetch(url);
        const json = await response.json();
        
        adminData = json.data || [];
        
        loading.classList.add('hidden');
        
        if (adminData.length === 0) {
            empty.classList.remove('hidden');
        } else {
            updateAdminStats();
            renderAdminList('');
        }
    } catch (error) {
        loading.classList.add('hidden');
        empty.classList.remove('hidden');
        console.error('Erro ao carregar dados admin:', error);
    }
}

function updateAdminStats() {
    document.getElementById('admin-total').textContent = adminData.length;
    
    const comFoto = adminData.filter(d => d.foto === 'Sim').length;
    document.getElementById('admin-com-foto').textContent = comFoto;
    
    const now = new Date();
    const sevenDays = 7 * 24 * 60 * 60 * 1000;
    const recentes = adminData.filter(d => {
        const ts = new Date(d.timestamp);
        return (now - ts) < sevenDays;
    }).length;
    document.getElementById('admin-recentes').textContent = recentes;
    
    // Popular sugestões de busca
    updateSearchSuggestions();
}

function updateSearchSuggestions() {
    const datalist = document.getElementById('admin-search-suggestions');
    if (!datalist) return;
    
    const suggestions = new Set();
    adminData.forEach(d => {
        if (d.nome) suggestions.add(d.nome);
        if (d.email) suggestions.add(d.email);
        if (d.cidade) suggestions.add(d.cidade);
        if (d.estado) suggestions.add(d.cidade + '/' + d.estado);
        if (d.instagram) suggestions.add(d.instagram.replace(/^@+/, ''));
        if (d.cursos) {
            d.cursos.split('\n').forEach(c => {
                const nome = c.replace(/\s*\(\d{4}\)/, '').trim();
                if (nome) suggestions.add(nome);
            });
        }
    });
    
    datalist.innerHTML = '';
    suggestions.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s;
        datalist.appendChild(opt);
    });
}

function renderAdminList(filter, dateFilteredData) {
    const list = document.getElementById('admin-list');
    const empty = document.getElementById('admin-empty');
    list.innerHTML = '';
    
    let baseData = dateFilteredData || getDateFilteredData();
    let filtered = baseData;
    
    if (filter) {
        const filterLower = filter.toLowerCase();
        filtered = baseData.filter(d => {
            const searchStr = [d.nome, d.email, d.cidade, d.cursos, d.instagram].join(' ').toLowerCase();
            return searchStr.includes(filterLower);
        });
    }
    
    if (filtered.length === 0) {
        empty.classList.remove('hidden');
        return;
    }
    
    empty.classList.add('hidden');
    
    // Ordenar mais recente primeiro
    filtered.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    filtered.forEach(item => {
        const card = document.createElement('div');
        card.className = 'admin-card';
        
        const initials = (item.nome || '?').charAt(0).toUpperCase();
        const date = formatAdminDate(item.timestamp);
        const cursos = (item.cursos || '').replace(/\n/g, ', ');
        
        let avatarContent = `<span>${initials}</span>`;
        let avatarClickable = false;
        if (item.foto_link && item.foto_link.indexOf('drive.google.com') !== -1) {
            const fileId = item.foto_link.match(/\/d\/([a-zA-Z0-9_-]+)/);
            if (fileId && fileId[1]) {
                avatarContent = `<img src="https://lh3.googleusercontent.com/d/${fileId[1]}=w100" alt="">`;
                avatarClickable = true;
            }
        }
        
        card.innerHTML = `
            <div class="admin-card-avatar">${avatarContent}</div>
            <div class="admin-card-info">
                <div class="admin-card-name">${item.nome || 'Sem nome'}</div>
                <div class="admin-card-detail">${cursos || 'Sem curso'}</div>
            </div>
            <span class="admin-card-date">${date}</span>
            <div class="admin-card-actions">
                <button class="admin-card-btn delete" data-row="${item._row}" data-name="${item.nome}" title="Excluir">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                    </svg>
                </button>
            </div>
        `;
        
        // Click no card abre detalhe
        card.addEventListener('click', (e) => {
            if (e.target.closest('.admin-card-btn')) return;
            showAdminDetail(item);
        });
        
        // Click no delete
        card.querySelector('.admin-card-btn.delete').addEventListener('click', (e) => {
            e.stopPropagation();
            adminDeleteRow = parseInt(e.currentTarget.dataset.row);
            adminDeleteName = e.currentTarget.dataset.name;
            document.getElementById('admin-delete-name').textContent = adminDeleteName;
            document.getElementById('admin-delete-confirm').classList.remove('hidden');
            pushAppState('admin-delete');
        });
        
        list.appendChild(card);
    });
}

function showAdminDetail(item) {
    const content = document.getElementById('admin-detail-content');
    const cursos = (item.cursos || '').replace(/\n/g, '<br>');
    
    let photoHtml = '<span style="font-size:2rem;color:var(--color-gray-200);">👤</span>';
    let photoFullUrl = '';
    if (item.foto_link && item.foto_link.indexOf('drive.google.com') !== -1) {
        const fileId = item.foto_link.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileId && fileId[1]) {
            const thumbUrl = `https://lh3.googleusercontent.com/d/${fileId[1]}=w200`;
            photoFullUrl = `https://lh3.googleusercontent.com/d/${fileId[1]}=w800`;
            photoHtml = `<img src="${thumbUrl}" alt="" style="cursor:pointer" onclick="openPhotoFullscreen('${photoFullUrl}')">`;
        }
    }
    
    const whatsappClean = (item.whatsapp || '').replace(/\D/g, '');
    
    content.innerHTML = `
        <div class="admin-detail-photo">${photoHtml}</div>
        <h2 class="admin-detail-name">${item.nome || 'Sem nome'}</h2>
        <p class="admin-detail-location">${item.cidade ? item.cidade + '/' + item.estado : ''}</p>
        
        <div class="admin-detail-section">
            <div class="admin-detail-label">Email</div>
            <div class="admin-detail-value"><a href="mailto:${item.email}">${item.email || '-'}</a></div>
        </div>
        
        <div class="admin-detail-section">
            <div class="admin-detail-label">WhatsApp</div>
            <div class="admin-detail-value">${item.whatsapp || '-'}</div>
        </div>
        
        <div class="admin-detail-section">
            <div class="admin-detail-label">Cursos</div>
            <div class="admin-detail-value">${cursos || '-'}</div>
        </div>
        
        ${item.experiencia ? `
        <div class="admin-detail-section">
            <div class="admin-detail-label">Experiência</div>
            <div class="admin-detail-value">${item.experiencia}</div>
        </div>` : ''}
        
        ${item.instagram ? `
        <div class="admin-detail-section">
            <div class="admin-detail-label">Instagram</div>
            <div class="admin-detail-value"><a href="https://instagram.com/${item.instagram.replace(/^@+/, '')}" target="_blank">@${item.instagram.replace(/^@+/, '')}</a></div>
        </div>` : ''}
        
        ${item.portfolio ? `
        <div class="admin-detail-section">
            <div class="admin-detail-label">Portfólio</div>
            <div class="admin-detail-value"><a href="${item.portfolio}" target="_blank">${item.portfolio}</a></div>
        </div>` : ''}
        
        ${item.linkedin ? `
        <div class="admin-detail-section">
            <div class="admin-detail-label">LinkedIn</div>
            <div class="admin-detail-value"><a href="${item.linkedin}" target="_blank">Ver perfil</a></div>
        </div>` : ''}
        
        ${item.sobre ? `
        <div class="admin-detail-section">
            <div class="admin-detail-label">Sobre</div>
            <div class="admin-detail-value">${item.sobre}</div>
        </div>` : ''}
        
        <div class="admin-detail-section">
            <div class="admin-detail-label">Cadastrado em</div>
            <div class="admin-detail-value">${formatAdminDate(item.timestamp, true)}</div>
        </div>
        
        <div class="admin-detail-actions">
            ${whatsappClean ? `
            <button class="admin-detail-btn whatsapp" onclick="window.open('https://wa.me/55${whatsappClean}', '_blank')">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347z"/><path d="M12 0C5.373 0 0 5.373 0 12c0 2.625.846 5.059 2.284 7.034L.789 23.492l4.637-1.467A11.932 11.932 0 0012 24c6.627 0 12-5.373 12-12S18.627 0 12 0zm0 21.75c-2.171 0-4.178-.69-5.822-1.863l-.417-.283-2.751.87.914-2.669-.31-.447A9.714 9.714 0 012.25 12C2.25 6.624 6.624 2.25 12 2.25S21.75 6.624 21.75 12 17.376 21.75 12 21.75z"/></svg>
                WhatsApp
            </button>` : ''}
            <button class="admin-detail-btn delete" onclick="adminDeleteFromDetail(${item._row}, '${(item.nome || '').replace(/'/g, "\\'")}')">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
                </svg>
                Excluir
            </button>
        </div>
    `;
    
    document.getElementById('admin-detail').classList.remove('hidden');
    pushAppState('admin-detail');
}

function adminDeleteFromDetail(row, name) {
    document.getElementById('admin-detail').classList.add('hidden');
    adminDeleteRow = row;
    adminDeleteName = name;
    document.getElementById('admin-delete-name').textContent = name;
    document.getElementById('admin-delete-confirm').classList.remove('hidden');
    pushAppState('admin-delete');
}

async function adminDeleteConfirmed() {
    const confirmOverlay = document.getElementById('admin-delete-confirm');
    const deleteBtn = document.getElementById('admin-delete-yes');
    
    deleteBtn.textContent = 'Excluindo...';
    deleteBtn.disabled = true;
    
    try {
        const url = CONFIG.GOOGLE_SCRIPT_URL + '?action=delete&' + getAdminAuthParam() + '&row=' + adminDeleteRow;
        const response = await fetch(url);
        const json = await response.json();
        
        if (json.success) {
            confirmOverlay.classList.add('hidden');
            loadAdminData(); // Recarregar lista
        } else {
            alert('Erro ao excluir: ' + (json.error || 'Tente novamente'));
        }
    } catch (error) {
        alert('Erro de conexão. Tente novamente.');
    } finally {
        deleteBtn.textContent = 'Excluir';
        deleteBtn.disabled = false;
    }
}

function formatAdminDate(timestamp, full) {
    if (!timestamp) return '';
    const d = new Date(timestamp);
    if (isNaN(d)) return '';
    
    const day = d.getDate().toString().padStart(2, '0');
    const month = (d.getMonth() + 1).toString().padStart(2, '0');
    const year = d.getFullYear();
    
    if (full) {
        const hours = d.getHours().toString().padStart(2, '0');
        const mins = d.getMinutes().toString().padStart(2, '0');
        return `${day}/${month}/${year} às ${hours}:${mins}`;
    }
    
    return `${day}/${month}`;
}

// ========================================
// Admin Change Password
// ========================================

async function adminChangePassword() {
    const currentInput = document.getElementById('admin-pw-current');
    const newInput = document.getElementById('admin-pw-new');
    const confirmInput = document.getElementById('admin-pw-confirm');
    const errorEl = document.getElementById('admin-pw-error');
    const successEl = document.getElementById('admin-pw-success');
    const btn = document.getElementById('admin-pw-btn');
    
    errorEl.classList.add('hidden');
    successEl.classList.add('hidden');
    
    const current = currentInput.value.trim();
    const newPin = newInput.value.trim();
    const confirm = confirmInput.value.trim();
    
    // Validações locais básicas
    if (!current) {
        errorEl.textContent = 'Digite a senha atual';
        errorEl.classList.remove('hidden');
        return;
    }
    
    if (newPin.length < 4) {
        errorEl.textContent = 'Nova senha deve ter pelo menos 4 caracteres';
        errorEl.classList.remove('hidden');
        return;
    }
    
    if (newPin !== confirm) {
        errorEl.textContent = 'As senhas não coincidem';
        errorEl.classList.remove('hidden');
        return;
    }
    
    if (newPin === current) {
        errorEl.textContent = 'Nova senha deve ser diferente da atual';
        errorEl.classList.remove('hidden');
        return;
    }
    
    btn.textContent = 'Alterando...';
    btn.disabled = true;
    
    try {
        const url = CONFIG.GOOGLE_SCRIPT_URL + '?action=changepin&current=' + encodeURIComponent(current) + '&newpin=' + encodeURIComponent(newPin);
        const response = await fetch(url);
        const json = await response.json();
        
        if (json.success) {
            ADMIN_PIN = newPin;
            successEl.classList.remove('hidden');
            currentInput.value = '';
            newInput.value = '';
            confirmInput.value = '';
            
            setTimeout(() => {
                document.getElementById('admin-change-pw').classList.add('hidden');
            }, 1500);
        } else {
            errorEl.textContent = json.error || 'Erro ao alterar senha';
            errorEl.classList.remove('hidden');
        }
    } catch (error) {
        errorEl.textContent = 'Erro de conexão. Tente novamente.';
        errorEl.classList.remove('hidden');
    } finally {
        btn.textContent = 'Alterar Senha';
        btn.disabled = false;
    }
}

// ========================================
// Biometric Authentication (WebAuthn)
// ========================================

const BIOMETRIC_STORAGE_KEY = 'emda_biometric_credential';

function initBiometric() {
    // Verificar se WebAuthn é suportado
    if (!window.PublicKeyCredential) {
        return;
    }
    
    // Verificar se já tem credencial registrada neste dispositivo
    const hasCredential = localStorage.getItem(BIOMETRIC_STORAGE_KEY);
    
    if (hasCredential) {
        // Mostrar botão de login biométrico
        showBiometricLoginBtn();
    }
}

function showBiometricLoginBtn() {
    const divider = document.getElementById('biometric-divider');
    const btn = document.getElementById('biometric-login-btn');
    if (divider) divider.classList.remove('hidden');
    if (btn) {
        btn.classList.remove('hidden');
        btn.addEventListener('click', biometricLogin);
    }
}

function showBiometricRegisterBtn() {
    const section = document.getElementById('biometric-register-section');
    if (!section) return;
    
    // Só mostra se WebAuthn suportado e não tem credencial ainda
    if (!window.PublicKeyCredential) return;
    
    const hasCredential = localStorage.getItem(BIOMETRIC_STORAGE_KEY);
    if (hasCredential) return;
    
    section.classList.remove('hidden');
    
    const btn = document.getElementById('biometric-register-btn');
    // Remover listener antigo
    const newBtn = btn.cloneNode(true);
    btn.parentNode.replaceChild(newBtn, btn);
    
    newBtn.addEventListener('click', biometricRegister);
}

async function biometricRegister() {
    try {
        // Gerar um ID único para esta credencial
        const userId = new Uint8Array(16);
        crypto.getRandomValues(userId);
        
        const challenge = new Uint8Array(32);
        crypto.getRandomValues(challenge);
        
        const createOptions = {
            publicKey: {
                rp: {
                    name: 'EMDA Conexão Moda',
                    id: window.location.hostname
                },
                user: {
                    id: userId,
                    name: 'admin@emda',
                    displayName: 'Administrador EMDA'
                },
                challenge: challenge,
                pubKeyCredParams: [
                    { type: 'public-key', alg: -7 },   // ES256
                    { type: 'public-key', alg: -257 }  // RS256
                ],
                authenticatorSelection: {
                    authenticatorAttachment: 'platform',
                    userVerification: 'required',
                    residentKey: 'preferred'
                },
                timeout: 60000,
                attestation: 'none'
            }
        };
        
        const credential = await navigator.credentials.create(createOptions);
        
        // Salvar ID da credencial no localStorage
        const credentialId = btoa(String.fromCharCode(...new Uint8Array(credential.rawId)));
        localStorage.setItem(BIOMETRIC_STORAGE_KEY, credentialId);
        
        alert('Biometria ativada com sucesso neste dispositivo!');
        
        // Esconder botão de registro e mostrar botão de login
        document.getElementById('biometric-register-section').classList.add('hidden');
        showBiometricLoginBtn();
        
    } catch (error) {
        if (error.name === 'NotAllowedError') {
            alert('Biometria cancelada. Tente novamente quando quiser.');
        } else {
            console.error('Erro ao registrar biometria:', error);
            alert('Não foi possível ativar a biometria neste dispositivo.');
        }
    }
}

async function biometricLogin() {
    try {
        const credentialId = localStorage.getItem(BIOMETRIC_STORAGE_KEY);
        if (!credentialId) {
            alert('Nenhuma biometria cadastrada neste dispositivo.');
            return;
        }
        
        const challenge = new Uint8Array(32);
        crypto.getRandomValues(challenge);
        
        const rawId = Uint8Array.from(atob(credentialId), c => c.charCodeAt(0));
        
        const getOptions = {
            publicKey: {
                challenge: challenge,
                allowCredentials: [{
                    type: 'public-key',
                    id: rawId,
                    transports: ['internal']
                }],
                userVerification: 'required',
                timeout: 60000
            }
        };
        
        await navigator.credentials.get(getOptions);
        
        // Biometria validada — abrir painel admin
        adminAuthMode = 'bio';
        document.getElementById('admin-login').classList.add('hidden');
        openAdminPanel();
        
    } catch (error) {
        if (error.name === 'NotAllowedError') {
            // Usuário cancelou
        } else {
            console.error('Erro na biometria:', error);
            alert('Falha na autenticação biométrica. Use a senha.');
        }
    }
}

// ========================================
// Biometric Settings Toggle
// ========================================

function updateBiometricSettings() {
    const section = document.getElementById('biometric-settings');
    const toggleBtn = document.getElementById('biometric-toggle-btn');
    const statusText = document.getElementById('biometric-status-text');
    
    if (!section || !window.PublicKeyCredential) return;
    
    section.classList.remove('hidden');
    
    const hasCredential = localStorage.getItem(BIOMETRIC_STORAGE_KEY);
    
    if (hasCredential) {
        toggleBtn.classList.add('active');
        statusText.textContent = 'Ativada neste dispositivo';
    } else {
        toggleBtn.classList.remove('active');
        statusText.textContent = 'Desativada neste dispositivo';
    }
    
    // Remover listener antigo
    const newBtn = toggleBtn.cloneNode(true);
    toggleBtn.parentNode.replaceChild(newBtn, toggleBtn);
    
    newBtn.addEventListener('click', async () => {
        const isActive = newBtn.classList.contains('active');
        
        if (isActive) {
            // Desativar
            localStorage.removeItem(BIOMETRIC_STORAGE_KEY);
            newBtn.classList.remove('active');
            statusText.textContent = 'Desativada neste dispositivo';
            // Esconder botão de biometria no login
            const bioBtn = document.getElementById('biometric-login-btn');
            const bioDivider = document.getElementById('biometric-divider');
            if (bioBtn) bioBtn.classList.add('hidden');
            if (bioDivider) bioDivider.classList.add('hidden');
        } else {
            // Ativar - registrar biometria
            await biometricRegister();
            updateBiometricSettings();
        }
    });
}

// ========================================
// Admin Photo Fullscreen
// ========================================

function openPhotoFullscreen(photoUrl) {
    const overlay = document.getElementById('admin-photo-full');
    const img = document.getElementById('admin-photo-full-img');
    const closeBtn = document.getElementById('admin-photo-full-close');
    if (!overlay || !img) return;
    
    img.src = photoUrl;
    overlay.classList.remove('hidden');
    pushAppState('admin-photo');
    
    // Reset zoom state
    let scale = 1;
    let posX = 0;
    let posY = 0;
    let lastTap = 0;
    let pinchStartDist = 0;
    let pinchStartScale = 1;
    let dragStartX = 0;
    let dragStartY = 0;
    let isDragging = false;
    
    function updateTransform() {
        img.style.transform = `translate(${posX}px, ${posY}px) scale(${scale})`;
        img.classList.toggle('zoomed', scale > 1);
    }
    
    function resetZoom() {
        scale = 1;
        posX = 0;
        posY = 0;
        updateTransform();
        img.classList.remove('zoomed', 'dragging');
    }
    
    function closeFullscreen() {
        resetZoom();
        overlay.classList.add('hidden');
    }
    
    // Close button
    closeBtn.onclick = closeFullscreen;
    
    // Double tap to zoom
    img.addEventListener('click', function handler(e) {
        const now = Date.now();
        if (now - lastTap < 300) {
            // Double tap
            e.preventDefault();
            if (scale > 1) {
                resetZoom();
            } else {
                scale = 3;
                // Zoom para o ponto do toque
                const rect = img.getBoundingClientRect();
                const x = e.clientX - rect.left - rect.width / 2;
                const y = e.clientY - rect.top - rect.height / 2;
                posX = -x * 2;
                posY = -y * 2;
                updateTransform();
            }
            lastTap = 0;
        } else {
            lastTap = now;
            // Single tap — fechar se não está com zoom
            setTimeout(() => {
                if (lastTap !== 0 && scale <= 1) {
                    closeFullscreen();
                }
                lastTap = 0;
            }, 300);
        }
    });
    
    // Pinch zoom
    img.addEventListener('touchstart', function(e) {
        if (e.touches.length === 2) {
            e.preventDefault();
            pinchStartDist = Math.hypot(
                e.touches[0].clientX - e.touches[1].clientX,
                e.touches[0].clientY - e.touches[1].clientY
            );
            pinchStartScale = scale;
        } else if (e.touches.length === 1 && scale > 1) {
            isDragging = true;
            dragStartX = e.touches[0].clientX - posX;
            dragStartY = e.touches[0].clientY - posY;
            img.classList.add('dragging');
        }
    }, { passive: false });
    
    img.addEventListener('touchmove', function(e) {
        if (e.touches.length === 2) {
            e.preventDefault();
            const dist = Math.hypot(
                e.touches[0].clientX - e.touches[1].clientX,
                e.touches[0].clientY - e.touches[1].clientY
            );
            scale = Math.min(Math.max(pinchStartScale * (dist / pinchStartDist), 1), 5);
            if (scale <= 1) { posX = 0; posY = 0; }
            updateTransform();
        } else if (e.touches.length === 1 && isDragging && scale > 1) {
            e.preventDefault();
            posX = e.touches[0].clientX - dragStartX;
            posY = e.touches[0].clientY - dragStartY;
            updateTransform();
        }
    }, { passive: false });
    
    img.addEventListener('touchend', function(e) {
        isDragging = false;
        img.classList.remove('dragging');
        if (scale <= 1.05) {
            resetZoom();
        }
    });
}

// ========================================
// Password Toggle (show/hide)
// ========================================

function initPasswordToggles() {
    document.querySelectorAll('.password-toggle-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.target;
            const input = document.getElementById(targetId);
            if (!input) return;
            
            const eyeOpen = btn.querySelector('.eye-open');
            const eyeClosed = btn.querySelector('.eye-closed');
            
            if (input.type === 'password') {
                input.type = 'text';
                eyeOpen.classList.add('hidden');
                eyeClosed.classList.remove('hidden');
            } else {
                input.type = 'password';
                eyeOpen.classList.remove('hidden');
                eyeClosed.classList.add('hidden');
            }
        });
    });
}

// ========================================
// Date Filter
// ========================================

let activeDateFilter = { type: 'days', days: 7 };

function initDateFilter() {
    const statBtn = document.getElementById('admin-stat-recentes');
    const filterPanel = document.getElementById('admin-date-filter');
    if (!statBtn || !filterPanel) return;
    
    // Toggle dropdown
    statBtn.addEventListener('click', () => {
        const isOpen = !filterPanel.classList.contains('hidden');
        filterPanel.classList.toggle('hidden');
        statBtn.classList.toggle('open');
    });
    
    // Quick filter buttons
    filterPanel.querySelectorAll('.date-filter-opt').forEach(btn => {
        btn.addEventListener('click', () => {
            // Update active state
            filterPanel.querySelectorAll('.date-filter-opt').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            const days = parseInt(btn.dataset.days);
            activeDateFilter = { type: 'days', days: days };
            
            // Update label
            updateDateFilterLabel(btn.textContent.trim());
            
            // Apply filter
            applyDateFilter();
            
            // Clear custom inputs
            document.getElementById('admin-date-from').value = '';
            document.getElementById('admin-date-to').value = '';
        });
    });
    
    // Custom date apply
    document.getElementById('admin-date-apply').addEventListener('click', () => {
        const from = document.getElementById('admin-date-from').value;
        const to = document.getElementById('admin-date-to').value;
        
        if (!from && !to) return;
        
        activeDateFilter = {
            type: 'custom',
            from: from ? new Date(from + 'T00:00:00') : null,
            to: to ? new Date(to + 'T23:59:59') : null
        };
        
        // Clear quick filter active
        filterPanel.querySelectorAll('.date-filter-opt').forEach(b => b.classList.remove('active'));
        
        // Update label
        let label = '';
        if (from && to) {
            label = formatShortDate(from) + ' - ' + formatShortDate(to);
        } else if (from) {
            label = 'A partir de ' + formatShortDate(from);
        } else {
            label = 'Até ' + formatShortDate(to);
        }
        updateDateFilterLabel(label);
        
        applyDateFilter();
    });
}

function formatShortDate(dateStr) {
    const parts = dateStr.split('-');
    return parts[2] + '/' + parts[1];
}

function updateDateFilterLabel(text) {
    const label = document.querySelector('#admin-stat-recentes .admin-stat-label');
    if (label) {
        label.innerHTML = text + ' <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>';
    }
}

function applyDateFilter() {
    const now = new Date();
    let filtered;
    
    if (activeDateFilter.type === 'days') {
        if (activeDateFilter.days === 0) {
            filtered = adminData;
        } else {
            const ms = activeDateFilter.days * 24 * 60 * 60 * 1000;
            filtered = adminData.filter(d => (now - new Date(d.timestamp)) < ms);
        }
    } else {
        filtered = adminData.filter(d => {
            const ts = new Date(d.timestamp);
            if (activeDateFilter.from && ts < activeDateFilter.from) return false;
            if (activeDateFilter.to && ts > activeDateFilter.to) return false;
            return true;
        });
    }
    
    // Update counter
    document.getElementById('admin-recentes').textContent = filtered.length;
    
    // Re-render list with date filter
    renderAdminList(document.getElementById('admin-search-input').value, filtered);
}

function getDateFilteredData() {
    const now = new Date();
    
    if (activeDateFilter.type === 'days') {
        if (activeDateFilter.days === 0) return adminData;
        const ms = activeDateFilter.days * 24 * 60 * 60 * 1000;
        return adminData.filter(d => (now - new Date(d.timestamp)) < ms);
    } else {
        return adminData.filter(d => {
            const ts = new Date(d.timestamp);
            if (activeDateFilter.from && ts < activeDateFilter.from) return false;
            if (activeDateFilter.to && ts > activeDateFilter.to) return false;
            return true;
        });
    }
}
