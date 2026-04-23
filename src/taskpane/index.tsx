import React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

Office.onReady((info) => {
  console.log("[E-postrydder] Office.onReady:", info.host, info.platform);
  const container = document.getElementById("root");
  if (container) {
    try {
      const root = createRoot(container);
      root.render(<App />);
    } catch (e: any) {
      container.innerHTML = `<div style="padding:16px;font-family:sans-serif"><h3>Oppstartsfeil</h3><pre style="white-space:pre-wrap;color:#c00">${e.message}\n\n${e.stack || ""}</pre></div>`;
    }
  }
}).catch((err: any) => {
  const container = document.getElementById("root");
  if (container) {
    container.innerHTML = `<div style="padding:16px;font-family:sans-serif"><h3>Office.js-feil</h3><pre style="white-space:pre-wrap;color:#c00">${err.message || err}</pre></div>`;
  }
});
