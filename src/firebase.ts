import { initializeApp } from 'firebase/app';
import { getDatabase } from 'firebase/database';

const firebaseConfig = {
  apiKey: 'AIzaSyDo6SzLDWTrdChVRZiCvIbZP-RHTn6QM',
  authDomain: 'genmo-hc.firebaseapp.com',
  databaseURL: 'https://genmo-hc-default-rtdb.firebaseio.com',
  projectId: 'genmo-hc',
  storageBucket: 'genmo-hc.firebasestorage.app',
  messagingSenderId: '579105359163',
  appId: '1:579105359163:web:85530e91220daff0b8e31e',
  measurementId: 'G-C6PMJLLPDK',
};

const app = initializeApp(firebaseConfig);
export const db = getDatabase(app);
