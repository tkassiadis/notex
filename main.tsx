// src/main.tsx
import React from "react";
import ReactDOM from "react-dom/client";
import App from "./pages/App";
import "./index.css";

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// ── Registro do Service Worker ────────────────────────────────
// Registra apenas em produção (HTTPS) ou localhost
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker
      .register("/sw.js", { scope: "/" })
      .then((registration) => {
        console.log("[PWA] Service Worker registrado:", registration.scope);

        // Verifica atualizações a cada vez que o app abre
        registration.addEventListener("updatefound", () => {
          const newWorker = registration.installing;
          if (!newWorker) return;
          newWorker.addEventListener("statechange", () => {
            if (
              newWorker.state === "installed" &&
              navigator.serviceWorker.controller
            ) {
              // Nova versão disponível — notifica o usuário discretamente
              console.log("[PWA] Nova versão disponível. Recarregue para atualizar.");
            }
          });
        });
      })
      .catch((err) => {
        console.warn("[PWA] Service Worker falhou:", err);
      });
  });
}
