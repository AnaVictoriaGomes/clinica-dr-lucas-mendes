/**
 * Painel KOMMO – Conector Oficial Firebase Firestore + Firebase Auth
 * Clínica Dr. Lucas Mendes
 */

const firebaseConfig = {
  apiKey: "AIzaSyDzfLT0LTgFQYW-DWOPmpuu-kq9bYTjFSE",
  authDomain: "clinica-dr-lucas-mendes.firebaseapp.com",
  projectId: "clinica-dr-lucas-mendes",
  storageBucket: "clinica-dr-lucas-mendes.firebasestorage.app",
  messagingSenderId: "1079458841545",
  appId: "1:1079458841545:web:f987226f9e3b85c5a21e18",
  measurementId: "G-YLDFD3J8YY"
};

let app = null;
let db = null;
let auth = null;
let unsubWeeks = null;
let unsubAppts = null;

function stopListeners() {
  if (typeof unsubWeeks === 'function') {
    try { unsubWeeks(); } catch(e){}
    unsubWeeks = null;
  }
  if (typeof unsubAppts === 'function') {
    try { unsubAppts(); } catch(e){}
    unsubAppts = null;
  }
}

function startListeners() {
  stopListeners();
  if (!db) return;

  // Escuta semanas em tempo real apenas quando autenticado
  unsubWeeks = onSnapshot(collection(db, "weeks"), snap => {
    const w = {};
    snap.forEach(d => { w[d.id] = d.data(); });
    if (typeof window.setWeeks === 'function') {
      window.setWeeks(w);
    }
  }, err => {
    console.warn("Aviso ao escutar semanas no Firestore:", err);
  });

  // Escuta consultas em tempo real apenas quando autenticado
  unsubAppts = onSnapshot(collection(db, "appts"), snap => {
    const a = {};
    snap.forEach(d => { a[d.id] = d.data(); });
    if (typeof window.setAppts === 'function') {
      window.setAppts(a);
    }
  }, err => {
    console.warn("Aviso ao escutar consultas no Firestore:", err);
  });
}

async function initFirebase() {
  try {
    const { initializeApp } = await import("https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js");
    const { getFirestore, doc, setDoc, deleteDoc, onSnapshot, collection } = await import("https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js");
    const { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged, setPersistence, browserSessionPersistence } = await import("https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js");

    app = initializeApp(firebaseConfig);
    db = getFirestore(app);
    auth = getAuth(app);

    // Persistência na sessão da aba para segurança
    try {
      await setPersistence(auth, browserSessionPersistence);
    } catch(pErr) {
      console.warn("Persistência padrão mantida:", pErr);
    }

    window.DATA = {
      ready: true,
      authenticated: false,
      user: null,
      lastRead: Date.now(),

      async login(email, password) {
        if (!auth) throw new Error("Firebase Auth não carregado.");
        const cred = await signInWithEmailAndPassword(auth, email, password);
        return cred.user;
      },

      async logout() {
        stopListeners();
        if (typeof window.clearDataOnLogout === 'function') {
          window.clearDataOnLogout();
        }
        if (auth) {
          await signOut(auth);
        }
      },

      async write(col, id, data) {
        if (!db) throw new Error("Banco de dados não conectado.");
        const ref = doc(db, col, id);
        if (data === null) {
          await deleteDoc(ref);
        } else {
          await setDoc(ref, data, { merge: true });
        }
      },

      async saveCfg(patch) {
        if (!db) throw new Error("Banco de dados não conectado.");
        await setDoc(doc(db, 'config', 'general'), patch, { merge: true });
      },

      cardHtml: () => `
        <h2>Status do Firebase Firestore</h2>
        <p class="note" style="color:var(--g);font-weight:600;font-size:14px">● Conectado ao projeto: <code>${firebaseConfig.projectId}</code></p>
        <p class="note">Autenticação: ${window.DATA && window.DATA.user ? `<b>${window.DATA.user.email}</b> (UID: <code>${window.DATA.user.uid}</code>)` : 'Não autenticado'}</p>
        <p class="note">As gravações são persistidas no Firestore em nuvem e sincronizadas em tempo real entre abas e dispositivos.</p>
      `,
      migHtml: () => ''
    };

    // Monitoramento oficial do estado de autenticação
    onAuthStateChanged(auth, user => {
      if (user) {
        window.DATA.user = user;
        window.DATA.authenticated = true;
        window.DATA.lastRead = Date.now();
        console.log("Firebase Auth: usuário logado com sucesso:", user.email, user.uid);
        startListeners();
        if (typeof window.onAuthSuccess === 'function') {
          window.onAuthSuccess(user);
        }
      } else {
        window.DATA.user = null;
        window.DATA.authenticated = false;
        console.log("Firebase Auth: nenhum usuário autenticado");
        stopListeners();
        if (typeof window.onAuthRequired === 'function') {
          window.onAuthRequired();
        }
      }
      if (typeof window.stLine === 'function') window.stLine();
    });

    console.log("Firebase Firestore e Auth configurados:", firebaseConfig.projectId);
    if (typeof window.stLine === 'function') window.stLine();
    return;
  } catch (err) {
    console.error("Erro crítico ao inicializar o Firebase:", err);
  }

  // Fallback seguro em caso de indisponibilidade
  window.DATA = {
    ready: false,
    authenticated: false,
    user: null,
    lastRead: Date.now(),
    async login() { throw new Error("Conexão indisponível."); },
    async logout() {},
    async write() { throw new Error("Sem conexão com o Firebase."); },
    async saveCfg() { throw new Error("Sem conexão com o Firebase."); },
    cardHtml: () => `<h2>Status do Firebase</h2><p class="note" style="color:var(--r)">Falha ao conectar no projeto em nuvem.</p>`,
    migHtml: () => ''
  };
  if (typeof window.stLine === 'function') window.stLine();
}

initFirebase();
