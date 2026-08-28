import { defineConfig } from "vite";
import { tanstackStart } from "@tanstack/react-start/plugin/vite";
import netlify from "@netlify/vite-plugin-tanstack-start";
import viteReact from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";
import tsConfigPaths from "vite-tsconfig-paths";

// Configuração padrão do TanStack Start para Netlify.
// O plugin da Netlify gera corretamente as Functions/SSR para a plataforma,
// evitando o preset Cloudflare usado pela configuração do Lovable.
export default defineConfig({
  plugins: [
    tanstackStart(),
    netlify(),
    viteReact(),
    tailwindcss(),
    tsConfigPaths(),
  ],
  resolve: {
    dedupe: ["react", "react-dom", "@tanstack/react-router"],
  },
});
