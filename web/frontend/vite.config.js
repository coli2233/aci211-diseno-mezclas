/* Configuración de Vite (Equipo 6).
   - El plugin de React transpila nuestro JSX.
   - En desarrollo redirigimos las llamadas /api a la API local
     (FastAPI en 127.0.0.1:8000, la que arranca web/iniciar_web.bat).
   - "npm run build" genera dist/: ese es exactamente el paquete que
     Vercel construye y publica en producción. */
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: { "/api": "http://127.0.0.1:8000" },
  },
});
