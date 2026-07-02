// MSAL Configuration for Azure AD Authentication
const msalConfig = {
    auth: {
        clientId: '1c2381d0-cf26-4c4c-9caf-93d2435842c9',
        authority: 'https://login.microsoftonline.com/organizations',
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
    }
};

const loginRequest = {
    scopes: ['https://management.azure.com/.default']
};

let msalInstance = null;
let currentAccount = null;

function initMsal() {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    
    // Check if user is already logged in
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        currentAccount = accounts[0];
        return true;
    }
    return false;
}

async function login() {
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        currentAccount = response.account;
        return currentAccount;
    } catch (error) {
        console.error('Login failed:', error);
        throw error;
    }
}

async function logout() {
    try {
        await msalInstance.logoutPopup({
            account: currentAccount
        });
        currentAccount = null;
    } catch (error) {
        console.error('Logout failed:', error);
    }
}

async function getAccessToken() {
    if (!currentAccount) {
        throw new Error('No user logged in');
    }
    
    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: ['https://management.azure.com/.default'],
            account: currentAccount
        });
        return response.accessToken;
    } catch (error) {
        // If silent acquisition fails, try popup
        try {
            const response = await msalInstance.acquireTokenPopup({
                scopes: ['https://management.azure.com/.default'],
                account: currentAccount
            });
            return response.accessToken;
        } catch (popupError) {
            console.error('Token acquisition failed:', popupError);
            throw popupError;
        }
    }
}

function getCurrentUser() {
    return currentAccount;
}

function isLoggedIn() {
    return currentAccount !== null;
}
