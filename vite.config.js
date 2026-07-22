import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// ⚠️  IMPORTANTE: Cambia la base según dónde lo despliegas:
//
//   GitHub Pages (repo username.github.io)  → base: '/'
//   GitHub Pages (repo cualquier-nombre)    → base: '/nombre-del-repo/'
//   Vercel / Netlify / Render               → base: '/'
//
// Ejemplo para repo llamado "gestor-gastos": base: '/gestor-gastos/'

export default defineConfig({
  plugins: [react()],
  base: '/gestor-gastos/',   // ← cambia esto si tu repo tiene otro nombre
})
