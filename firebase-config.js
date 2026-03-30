// firebase-config.js
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";

// Your Firebase configuration (updated with your actual config)
const firebaseConfig = {
  apiKey: "AIzaSyBeSa9vEUV4PZgBNAl5pizqXpnbhBwUAoA",
  authDomain: "mambo-b70b0.firebaseapp.com",
  projectId: "mambo-b70b0",
  storageBucket: "mambo-b70b0.firebasestorage.app",
  messagingSenderId: "1064560518334",
  appId: "1:1064560518334:web:15a76cf31988615e6deaef",
  measurementId: "G-SC4EJ76344",
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

export { auth, db };
