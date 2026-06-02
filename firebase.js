// Firebase SDK
import { initializeApp } from "https://www.gstatic.com/firebasejs/12.2.1/firebase-app.js";

import {
    getFirestore,
    doc,
    setDoc,
    getDoc,
    getDocs,
    onSnapshot,
    collection
}
from "https://www.gstatic.com/firebasejs/12.2.1/firebase-firestore.js";

const firebaseConfig = {

    apiKey: "AIzaSy4W5NgSimObS9LScI-EUjPqP9_ai6h-ePU",

    authDomain:
        "ceremonia-lanzamiento.firebaseapp.com",

    projectId:
        "ceremonia-lanzamiento",

    storageBucket:
        "ceremonia-lanzamiento.firebasestorage.app",

    messagingSenderId:
        "291757329730",

    appId:
        "1:291757329730:web:3a4795b7755173b4fd4cdb"

};

const app =
    initializeApp(firebaseConfig);

const db =
    getFirestore(app);

// dejar disponible para script.js

window.db = db;
window.firebaseFirestore = {
    doc,
    setDoc,
    getDoc,
    getDocs,
    onSnapshot,
    collection
};

console.log("Firebase conectado");

window.coleccionIngresos =
    collection(
        db,
        "ingresos"
    );