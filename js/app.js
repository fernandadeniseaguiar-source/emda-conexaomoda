/* ========================================
   EMDA - Banco de Talentos
   Application JavaScript
======================================== */

// ========================================
// Configuration
// ========================================

const CONFIG = {
    // URL do Google Apps Script Web App (você precisará criar e colocar aqui)
    GOOGLE_SCRIPT_URL: 'YOUR_GOOGLE_APPS_SCRIPT_URL',
    
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

// ========================================
// Initialization
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    initSplashScreen();
    initWelcomeScreen();
    initPhotoUpload();
    initFormNavigation();
    initFormValidation();
    initPhoneMask();
    initCityAutocomplete();
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
        // Fade out welcome screen
        elements.welcomeScreen.style.opacity = '0';
        elements.welcomeScreen.style.transform = 'translateY(-20px)';
        elements.welcomeScreen.style.transition = 'all 0.5s ease';
        
        setTimeout(() => {
            elements.welcomeScreen.classList.add('hidden');
            elements.formContainer.classList.remove('hidden');
            
            // Scroll to top
            window.scrollTo({ top: 0, behavior: 'instant' });
        }, 500);
    });
}

// ========================================
// Photo Upload
// ========================================

function initPhotoUpload() {
    const photoInput = document.getElementById('photo-input');
    const cameraInput = document.getElementById('camera-input');
    const btnGallery = document.getElementById('btn-gallery');
    const btnSelfie = document.getElementById('btn-selfie');
    
    // Gallery button
    btnGallery.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        photoInput.click();
    });
    
    // Selfie button
    btnSelfie.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        cameraInput.click();
    });
    
    // Handle file selection
    photoInput.addEventListener('change', handlePhotoSelect);
    cameraInput.addEventListener('change', handlePhotoSelect);
}

function handlePhotoSelect(e) {
    const file = e.target.files[0];
    
    if (!file) return;
    
    // Validate file type
    if (!file.type.startsWith('image/')) {
        alert('Por favor, selecione uma imagem válida.');
        return;
    }
    
    // Validate file size (max 5MB)
    if (file.size > 5 * 1024 * 1024) {
        alert('A imagem deve ter no máximo 5MB.');
        return;
    }
    
    // Read and preview
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
// Form Navigation
// ========================================

function initFormNavigation() {
    // Next buttons
    document.querySelectorAll('.btn-next').forEach(btn => {
        btn.addEventListener('click', () => {
            const nextStep = parseInt(btn.dataset.next);
            if (validateStep(currentStep)) {
                goToStep(nextStep);
            }
        });
    });
    
    // Previous buttons
    document.querySelectorAll('.btn-prev').forEach(btn => {
        btn.addEventListener('click', () => {
            const prevStep = parseInt(btn.dataset.prev);
            goToStep(prevStep);
        });
    });
}

function goToStep(step) {
    // Update form steps
    elements.formSteps.forEach(formStep => {
        formStep.classList.remove('active');
        if (parseInt(formStep.dataset.step) === step) {
            formStep.classList.add('active');
        }
    });
    
    // Update progress indicators
    elements.progressSteps.forEach((progressStep, index) => {
        progressStep.classList.remove('active', 'completed');
        
        if (index + 1 < step) {
            progressStep.classList.add('completed');
        } else if (index + 1 === step) {
            progressStep.classList.add('active');
        }
    });
    
    // Update progress lines
    elements.progressFill.forEach((fill, index) => {
        if (index < step - 1) {
            fill.style.width = '100%';
        } else {
            fill.style.width = '0';
        }
    });
    
    currentStep = step;
    
    // Scroll to top of form
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
    
    // Step-specific validation
    if (step === 1) {
        // Email validation
        const emailInput = document.getElementById('email');
        if (emailInput.value && !isValidEmail(emailInput.value)) {
            isValid = false;
            highlightError(emailInput);
        }
        
        // Phone validation
        const phoneInput = document.getElementById('whatsapp');
        if (phoneInput.value && !isValidPhone(phoneInput.value)) {
            isValid = false;
            highlightError(phoneInput);
        }
    }
    
    if (step === 2) {
        // At least one course must be selected
        const cursos = document.querySelectorAll('input[name="cursos"]:checked');
        if (cursos.length === 0) {
            isValid = false;
            alert('Por favor, selecione pelo menos um curso.');
        }
    }
    
    if (!isValid) {
        // Shake animation for invalid form
        currentFormStep.style.animation = 'none';
        currentFormStep.offsetHeight; // Trigger reflow
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
    // Remove formatting
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
        
        // Click handler for suggestions
        suggestionsContainer.querySelectorAll('.autocomplete-item').forEach(item => {
            item.addEventListener('click', () => {
                cidadeInput.value = item.dataset.value.split(',')[0];
                suggestionsContainer.classList.remove('show');
                
                // Auto-select state
                const state = item.dataset.value.split(',')[1]?.trim();
                if (state) {
                    const estadoSelect = document.getElementById('estado');
                    if (estadoSelect) {
                        estadoSelect.value = state;
                    }
                }
            });
        });
    });
    
    // Keyboard navigation
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
    
    // Close on click outside
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
    
    // Validate final step
    if (!validateStep(3)) {
        return;
    }
    
    // Check terms
    const termos = document.getElementById('termos');
    if (!termos.checked) {
        alert('Por favor, aceite os termos para continuar.');
        return;
    }
    
    // Show loading state
    setLoadingState(true);
    
    // Collect form data
    const formData = collectFormData();
    
    try {
        // Send to Google Sheets
        await sendToGoogleSheets(formData);
        
        // Show success modal
        showSuccessModal();
        
    } catch (error) {
        console.error('Erro ao enviar:', error);
        alert('Ocorreu um erro ao enviar seu currículo. Por favor, tente novamente.');
    } finally {
        setLoadingState(false);
    }
}

function collectFormData() {
    // Get selected courses
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
    // Se a URL do Google Script não estiver configurada, simula sucesso
    if (CONFIG.GOOGLE_SCRIPT_URL === 'YOUR_GOOGLE_APPS_SCRIPT_URL') {
        console.log('Dados coletados (modo simulação):', data);
        // Simula delay de rede
        await new Promise(resolve => setTimeout(resolve, 1500));
        return { success: true };
    }
    
    const response = await fetch(CONFIG.GOOGLE_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
    });
    
    return { success: true };
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
    // Check if already installed
    if (window.matchMedia('(display-mode: standalone)').matches) {
        console.log('App já está instalado');
        return;
    }
    
    // Check if dismissed recently (24 hours)
    const dismissedAt = localStorage.getItem('installDismissed');
    if (dismissedAt && (Date.now() - parseInt(dismissedAt)) < 24 * 60 * 60 * 1000) {
        return;
    }
    
    // Listen for beforeinstallprompt (Android/Chrome)
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        
        // Show install prompt after a delay
        setTimeout(() => {
            showInstallPrompt();
        }, 3000);
    });
    
    // Handle install button click
    elements.installAccept.addEventListener('click', handleInstallClick);
    
    // Handle dismiss button click
    elements.installDismiss.addEventListener('click', hideInstallPrompt);
    
    // Check for iOS
    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
    const isInStandaloneMode = window.navigator.standalone === true;
    
    if (isIOS && !isInStandaloneMode) {
        // Show iOS-specific instructions after delay
        setTimeout(() => {
            showIOSInstallPrompt();
        }, 5000);
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
    if (!deferredPrompt) {
        // Fallback for browsers that don't support beforeinstallprompt
        showIOSInstallPrompt();
        return;
    }
    
    // Show the install prompt
    deferredPrompt.prompt();
    
    // Wait for the user's response
    const { outcome } = await deferredPrompt.userChoice;
    
    console.log('Resultado da instalação:', outcome);
    
    // Clear the deferred prompt
    deferredPrompt = null;
    
    // Hide the install banner
    hideInstallPrompt();
}

function showIOSInstallPrompt() {
    // Create iOS-specific modal
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
    
    // Close on backdrop click
    modal.addEventListener('click', (e) => {
        if (e.target === modal) {
            modal.remove();
        }
    });
    
    // Hide the regular install prompt if visible
    hideInstallPrompt();
}

// ========================================
// Shake Animation (CSS-in-JS fallback)
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
