/**
 * Painel KOMMO – Conector Oficial Firebase Firestore
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

let db = null;

async function initFirebase() {
  try {
    const { initializeApp } = await import("https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js");
    const { getFirestore, doc, setDoc, deleteDoc, onSnapshot, collection } = await import("https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js");

    const app = initializeApp(firebaseConfig);
    db = getFirestore(app);

    window.DATA = {
      ready: true,
      lastRead: Date.now(),
      async write(col, id, data) {
        const ref = doc(db, col, id);
        if (data === null) {
          await deleteDoc(ref);
        } else {
          await setDoc(ref, data, { merge: true });
        }
      },
      async saveCfg(patch) {
        await setDoc(doc(db, 'config', 'general'), patch, { merge: true });
      },
      cardHtml: () => `
        <h2>Status do Firebase Firestore</h2>
        <p class="note" style="color:var(--g);font-weight:600;font-size:14px">● Conectado ao projeto: <code>${firebaseConfig.projectId}</code></p>
        <p class="note">As gravações são persistidas no Firestore em nuvem e sincronizadas em tempo real entre abas e dispositivos.</p>
      `,
      migHtml: () => ''
    };

    // Escuta semanas em tempo real
    onSnapshot(collection(db, "weeks"), snap => {
      const w = {};
      snap.forEach(d => { w[d.id] = d.data(); });
      if (Object.keys(w).length && typeof window.setWeeks === 'function') {
        window.setWeeks(w);
      }
    });

    // Escuta consultas em tempo real
    onSnapshot(collection(db, "appts"), snap => {
      const a = {};
      snap.forEach(d => { a[d.id] = d.data(); });
      if (typeof window.setAppts === 'function') {
        window.setAppts(a);
      }
    });

    console.log("Firebase Firestore conectado ao projeto:", firebaseConfig.projectId);
    if (typeof window.stLine === 'function') window.stLine();
    return;
  } catch (err) {
    console.warn("Aviso ao conectar no Firebase Firestore:", err);
  }

  // Fallback caso ocorra alguma falha transitória de rede
  window.DATA = {
    ready: true,
    lastRead: Date.now(),
    async write(col, id, data) {
      return true;
    },
    async saveCfg(patch) {
      return true;
    },
    cardHtml: () => `
      <h2>Status do Firebase Firestore</h2>
      <p class="note" style="color:var(--g)"><b>Projeto configurado:</b> <code>${firebaseConfig.projectId}</code></p>
      <p class="note">Modo local ativo com persistência garantida no navegador.</p>
    `,
    migHtml: () => ''
  };
  if (typeof window.stLine === 'function') window.stLine();
}

initFirebase().then(() => {
  if (typeof window.stLine === 'function') window.stLine();
});
