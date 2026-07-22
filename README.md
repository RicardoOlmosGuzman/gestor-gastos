# 💸 GestorGastos

Aplicación web para organizar, analizar y planificar gastos personales — hecha para la realidad chilena.

## ✨ Funciones

- 📍 **Gestión por dirección** — organiza gastos de distintas casas o lugares
- 📅 **Mes a mes** — categorías, ítems y sub-ítems con formato CLP automático  
- 📊 **Análisis** — gráficos de barras, torta y líneas para 1-12 meses
- 📥 **Exportar** — Excel y PDF directamente desde el navegador
- 🧠 **Planificación IA** — planes de ahorro, diagnóstico y más con Claude
- 🌙 **Modo oscuro** — incluido por defecto
- 💾 **Datos locales** — todo se guarda en tu navegador (localStorage)

## 🚀 Deploy rápido

### Opción A — GitHub Pages (gratis, URL propia)

1. **Fork o sube este repo a GitHub**
2. En `vite.config.js`, cambia `base` al nombre de tu repo:
   ```js
   base: '/nombre-de-tu-repo/',
   ```
3. En tu repo en GitHub → **Settings → Pages → Source → GitHub Actions**
4. Haz un `git push` y en ~2 minutos tendrás tu URL:
   ```
   https://tu-usuario.github.io/nombre-de-tu-repo/
   ```

### Opción B — Vercel (más fácil, 1 clic)

1. Sube el repo a GitHub
2. Entra a [vercel.com](https://vercel.com) → **New Project → Import**
3. Selecciona el repo → **Deploy**
4. En `vite.config.js` cambia `base: '/'`

### Opción C — Netlify Drop (sin cuenta)

1. Primero ejecuta localmente:
   ```bash
   npm install
   npm run build
   ```
2. Arrastra la carpeta `dist/` a [netlify.com/drop](https://app.netlify.com/drop)
3. ¡Listo! Te da URL al instante

## 🖥️ Desarrollo local

```bash
npm install
npm run dev
```

Abre http://localhost:5173

## 🔑 API Key para la IA

La función de Planificación IA usa Claude (Anthropic). Necesitas una API key:

1. Crea una cuenta en [console.anthropic.com](https://console.anthropic.com)
2. Ve a **API Keys → Create Key**
3. En la app, ve a **Plan IA → pega tu key**

La key se guarda solo en tu navegador (localStorage), no se sube a ningún servidor.

## 🛠️ Stack

- **React 18** + Vite
- **recharts** — gráficos
- **xlsx** — exportación a Excel
- **lucide-react** — íconos

---

Hecho con ❤️ en Chile 🇨🇱
