/**
 * ═══════════════════════════════════════════════════════════════
 * CITRO — Autenticación Microsoft (OPTIMIZADO)
 * Azure Active Directory con MSAL 2.x
 * ═══════════════════════════════════════════════════════════════
 */

// Estado global del usuario
const userState = {
    isLoggedIn: false,
    isAdmin: false,
    account: null,
    token: null,
    tokenExpiry: null,
    profile: null
};

let msalApp = null;
let tokenCache = new Map(); // Cache de tokens por scope

/**
 * Inicializar MSAL al cargar la página
 */
function initMSAL() {
    const config = {
        auth: {
            clientId: CONFIG.azure.clientId,
            authority: `https://login.microsoftonline.com/${CONFIG.azure.tenantId}`,
            redirectUri: window.location.origin + window.location.pathname,
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        },
        cache: {
            cacheLocation: 'sessionStorage',
            storeAuthStateInCookie: false
        }
    };

    msalApp = new msal.PublicClientApplication(config);

    // Manejar respuesta de redirect
    msalApp.handleRedirectPromise()
        .then(response => {
            if (response) {
                handleLoginSuccess(response);
            } else {
                tryRestoreSession();
            }
        })
        .catch(error => {
            console.error('Error MSAL:', error);
            if (CONFIG.options.debug) {
                alert('Error al inicializar autenticación: ' + error.message);
            }
        });
}

/**
 * Intentar restaurar sesión previa
 */
function tryRestoreSession() {
    const accounts = msalApp.getAllAccounts();
    
    if (accounts.length > 0) {
        msalApp.setActiveAccount(accounts[0]);
        loadProfile(accounts[0]).catch(() => {
            // Si falla, mostrar login
            showLoginUI();
        });
    } else {
        showLoginUI();
    }
}

/**
 * Mostrar UI de login
 */
function showLoginUI() {
    document.getElementById('login-container').style.display = 'block';
    document.getElementById('user-menu').style.display = 'none';
    document.getElementById('login-required-banner').style.display = 'flex';
}

/**
 * Iniciar sesión con Microsoft
 */
async function signInWithMicrosoft() {
    try {
        const loginRequest = {
            scopes: CONFIG.azure.scopes,
            prompt: 'select_account'
        };

        const response = await msalApp.loginPopup(loginRequest);
        await handleLoginSuccess(response);
        
    } catch (error) {
        if (error.errorCode === 'user_cancelled') {
            console.log('Login cancelado por el usuario');
        } else {
            console.error('Error en login:', error);
            alert('Error al iniciar sesión. Por favor, inténtelo nuevamente.\n\n' + error.message);
        }
    }
}

/**
 * Procesar login exitoso
 */
async function handleLoginSuccess(response) {
    msalApp.setActiveAccount(response.account);
    await loadProfile(response.account);
}

/**
 * Cargar perfil del usuario desde Microsoft Graph
 */
async function loadProfile(account) {
    try {
        // Obtener token
        const token = await getTokenSilently();
        userState.token = token;
        userState.tokenExpiry = Date.now() + (3600 * 1000); // 1 hora

        // Llamar a Graph API para obtener perfil
        const graphResponse = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        });

        if (!graphResponse.ok) {
            throw new Error('Error al obtener perfil de usuario');
        }

        const userData = await graphResponse.json();

        // Construir perfil
        userState.isLoggedIn = true;
        userState.account = account;
        userState.profile = {
            id: userData.id,
            nombre: userData.displayName,
            givenName: userData.givenName || userData.displayName.split(' ')[0],
            surname: userData.surname || '',
            email: userData.mail || userData.userPrincipalName,
            jobTitle: userData.jobTitle || '',
            department: userData.department || '',
            initials: getInitials(userData.displayName)
        };

        // Verificar dominio institucional si está configurado
        if (CONFIG.options.soloEmailUV) {
            const emailDomain = userState.profile.email.split('@')[1];
            if (emailDomain !== CONFIG.options.dominioPermitido) {
                alert(`⚠️ Solo se permite acceso con cuentas @${CONFIG.options.dominioPermitido}\n\nTu cuenta: ${userState.profile.email}`);
                await signOutMicrosoft();
                return;
            }
        }

        // Verificar si es administrador
        userState.isAdmin = CONFIG.admins
            .map(email => email.toLowerCase())
            .includes(userState.profile.email.toLowerCase());

        // Actualizar interfaz
        updateUIForLoggedIn();

        if (CONFIG.options.debug) {
            console.log('✅ Usuario autenticado:', userState.profile.nombre);
            console.log('📧 Email:', userState.profile.email);
            console.log('👑 Admin:', userState.isAdmin);
        }

    } catch (error) {
        console.error('Error al cargar perfil:', error);
        showLoginUI();
        throw error;
    }
}

/**
 * Obtener token silenciosamente (con cache)
 */
async function getTokenSilently() {
    // Verificar cache
    if (userState.token && userState.tokenExpiry && Date.now() < userState.tokenExpiry) {
        return userState.token;
    }

    const accounts = msalApp.getAllAccounts();
    if (!accounts.length) {
        throw new Error('No hay cuenta activa');
    }

    try {
        const response = await msalApp.acquireTokenSilent({
            scopes: CONFIG.azure.scopes,
            account: accounts[0]
        });
        
        userState.token = response.accessToken;
        userState.tokenExpiry = Date.now() + (3600 * 1000);
        
        return response.accessToken;
        
    } catch (error) {
        // Si falla silencioso, intentar con popup
        if (CONFIG.options.debug) {
            console.log('Token expirado, renovando con popup...');
        }
        
        const response = await msalApp.acquireTokenPopup({
            scopes: CONFIG.azure.scopes
        });
        
        userState.token = response.accessToken;
        userState.tokenExpiry = Date.now() + (3600 * 1000);
        
        return response.accessToken;
    }
}

/**
 * Obtener token de acceso (función pública)
 */
async function getAccessToken() {
    if (!userState.isLoggedIn) {
        throw new Error('Usuario no autenticado');
    }
    return getTokenSilently();
}

/**
 * Cerrar sesión
 */
async function signOutMicrosoft() {
    // Limpiar estado
    Object.assign(userState, {
        isLoggedIn: false,
        isAdmin: false,
        account: null,
        token: null,
        tokenExpiry: null,
        profile: null
    });

    // Limpiar cache
    tokenCache.clear();

    // Actualizar UI
    updateUIForLoggedOut();
    goToHome();

    // Logout de MSAL
    try {
        await msalApp.logoutPopup({
            postLogoutRedirectUri: window.location.origin + window.location.pathname
        });
    } catch (error) {
        // Si falla el popup, limpiar cache manualmente
        msalApp.clearCache();
    }
}

/**
 * Actualizar interfaz - usuario logueado
 */
function updateUIForLoggedIn() {
    const profile = userState.profile;

    // Ocultar login, mostrar menú de usuario
    document.getElementById('login-container').style.display = 'none';
    document.getElementById('login-required-banner').style.display = 'none';
    document.getElementById('user-menu').style.display = 'flex';

    // Actualizar datos del usuario
    document.getElementById('user-name-short').textContent = profile.givenName;
    document.getElementById('user-name-full').textContent = profile.nombre;
    document.getElementById('user-email-dd').textContent = profile.email;
    
    // Iniciales
    document.getElementById('user-initials').textContent = profile.initials;
    document.getElementById('user-avatar-large').textContent = profile.initials;

    // Mostrar panel de admin si corresponde
    if (userState.isAdmin) {
        document.getElementById('admin-panel-btn').style.display = 'block';
        if (CONFIG.options.debug) {
            console.log('👑 Acceso de administrador habilitado');
        }
    }
}

/**
 * Actualizar interfaz - usuario no logueado
 */
function updateUIForLoggedOut() {
    document.getElementById('login-container').style.display = 'block';
    document.getElementById('user-menu').style.display = 'none';
    document.getElementById('login-required-banner').style.display = 'flex';
    
    // Ocultar panel de admin
    const adminBtn = document.getElementById('admin-panel-btn');
    if (adminBtn) adminBtn.style.display = 'none';
}

/**
 * Toggle del menú de usuario
 */
function toggleUserDropdown() {
    const dropdown = document.getElementById('user-dropdown');
    dropdown.classList.toggle('show');
}

// Cerrar dropdown al hacer click fuera
document.addEventListener('click', (e) => {
    const userMenu = document.getElementById('user-menu');
    const dropdown = document.getElementById('user-dropdown');
    
    if (userMenu && dropdown && !userMenu.contains(e.target)) {
        dropdown.classList.remove('show');
    }
});

/**
 * Navegar a "Mis Solicitudes" (requiere login)
 */
function goToMisSolicitudes() {
    if (!userState.isLoggedIn) {
        signInWithMicrosoft();
        return;
    }
    showSection('mis-solicitudes');
    loadUserSolicitudes();
}

/**
 * Navegar a "Panel de Administración" (requiere admin)
 */
function goToAdminPanel() {
    if (!userState.isAdmin) {
        alert('⛔ Acceso denegado\n\nSolo administradores pueden acceder al panel.');
        return;
    }
    showSection('admin-panel');
    loadAdminData();
}

/**
 * Obtener iniciales del nombre
 */
function getInitials(name) {
    if (!name) return 'UV';
    
    const parts = name.trim().split(' ').filter(Boolean);
    
    if (parts.length === 1) {
        return parts[0].substring(0, 2).toUpperCase();
    }
    
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

/**
 * Verificar si el usuario está autenticado
 */
function isAuthenticated() {
    return userState.isLoggedIn;
}

/**
 * Verificar si el usuario es administrador
 */
function isAdmin() {
    return userState.isAdmin;
}

// Inicializar cuando carga el DOM
document.addEventListener('DOMContentLoaded', initMSAL);

// Log de versión
if (CONFIG.options.debug) {
    console.log(`🚀 ${CONFIG.version.nombre} v${CONFIG.version.numero}`);
    console.log(`📅 ${CONFIG.version.fecha}`);
}
