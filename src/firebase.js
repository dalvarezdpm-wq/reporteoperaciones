import { initializeApp } from "firebase/app";
import { getDatabase, ref, get, set, remove } from "firebase/database";

// Configuración real del proyecto Firebase "reporteaduana"
const firebaseConfig = {
  apiKey: "AIzaSyDWezJOllEOXroT0k0fRxDZcI3m4_Hfw_0",
  authDomain: "reporteaduana.firebaseapp.com",
  databaseURL: "https://reporteaduana-default-rtdb.firebaseio.com",
  projectId: "reporteaduana",
  storageBucket: "reporteaduana.firebasestorage.app",
  messagingSenderId: "139483644637",
  appId: "1:139483644637:web:904cf6b482e9c278805134",
};

const app = initializeApp(firebaseConfig);
export { app };
export const db = getDatabase(app);
export { ref, get, set, remove };
