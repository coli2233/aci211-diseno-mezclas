/* Punto de entrada (Equipo 6): montamos <App /> dentro de index.html.
   Vite compila este archivo y Vercel publica el resultado como web
   estática en aci211-diseno-mezclas.vercel.app. */
import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App.jsx";
import "./styles.css";

createRoot(document.getElementById("root")).render(<App />);
