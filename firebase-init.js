// ─── Firebase init (compat SDK, caricato da tutte le pagine) ──
// INCOLLA QUI la config del tuo progetto Firebase:
// Console Firebase → Impostazioni progetto → Le tue app → Config
const BM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyDL0g1mL_D8Msa-BTKWGj-BooJnAsyE29I",
  authDomain: "budget-manager-6207f.firebaseapp.com",
  projectId: "budget-manager-6207f",
  storageBucket: "budget-manager-6207f.firebasestorage.app",
  messagingSenderId: "349002704089",
  appId: "1:349002704089:web:f5710f7ca9a7f9c3144ac8"
};

// Account Google autorizzati ad accedere all'app.
// IMPORTANTE: deve combaciare con la lista in firestore.rules (da ripubblicare
// in console ogni volta che la modifichi).
const BM_ALLOWED_EMAILS = [
  'silvio.phy@gmail.com',
  'chiaraluca.mail@gmail.com',
];

firebase.initializeApp(BM_FIREBASE_CONFIG);

// App Check: verifica che le richieste a Firestore/Auth arrivino davvero da
// questa pagina (non da script/bot). Console Firebase → App Check → registra
// l'app web → provider reCAPTCHA v3 → incolla qui la site key.
const BM_RECAPTCHA_SITE_KEY = '6LeNmNAtAAAAAD2vog0fDnE9uXXaxYiyV1w3j6LA';
// L'attivazione inietta il badge reCAPTCHA nel <body>: va rimandata a quando
// il <body> esiste davvero, altrimenti fallisce (gli script sono in <head>).
function bmActivateAppCheck() {
  if (BM_RECAPTCHA_SITE_KEY) {
    firebase.appCheck().activate(BM_RECAPTCHA_SITE_KEY, true);
  }
}
if (document.body) bmActivateAppCheck();
else document.addEventListener('DOMContentLoaded', bmActivateAppCheck);

// Risolve con l'utente (o null) al primo stato auth noto
const bmAuthReady = new Promise(resolve => {
  const off = firebase.auth().onAuthStateChanged(u => { off(); resolve(u); });
});

// Multi-utente, ma solo per gli account in whitelist.
// Ogni utente autorizzato vede solo i propri dati (users/{uid}/…).
function bmIsAllowed(user) {
  return !!user && BM_ALLOWED_EMAILS.includes(user.email);
}
