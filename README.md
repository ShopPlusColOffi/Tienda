# 🛍️ Tienda Virtual — Lentes de Contacto

Sistema completo de tienda virtual con Google Apps Script como backend y Google Sheets como base de datos.

---

## 📁 Estructura de archivos

```
mi-tienda/
├── index.html      → Redirige a tienda.html
├── tienda.html     → Tienda pública (clientes)
├── admin.html      → Panel administrador
├── 404.html        → Redirige a tienda.html
├── Code.gs         → Backend (copiar a Apps Script)
└── README.md       → Este archivo
```

---

## 🚀 Configuración paso a paso

### Paso 1 — Crear el Google Sheets + Apps Script

1. Crea un nuevo Google Sheets en [sheets.google.com](https://sheets.google.com)
2. Ve a **Extensiones > Apps Script**
3. Borra el contenido del archivo `Code.gs` que aparece
4. Copia y pega TODO el contenido de `Code.gs` de este proyecto
5. Ejecuta la función `inicializarTienda()`:
   - En el menú del editor, selecciona `inicializarTienda` en el dropdown de funciones
   - Click en **▶ Ejecutar**
   - Acepta los permisos cuando te los pida
6. Verifica que se crearon las 5 hojas: `productos`, `categorias`, `pedidos`, `configuracion`, `zonas_domicilio`

### Paso 2 — Desplegar como Web App

1. En el editor de Apps Script, ve a **Implementar > Nueva implementación**
2. Tipo: **App web**
3. Configuración:
   - **Descripción**: Tienda Virtual API
   - **Ejecutar como**: **Yo** (tu cuenta)
   - **Quién tiene acceso**: **Cualquier persona**
4. Click en **Implementar**
5. **¡COPIA la URL que te da!** Se ve así:
   ```
   https://script.google.com/macros/s/AKfycbx.../exec
   ```

### Paso 3 — Configurar los archivos HTML

1. Abre `tienda.html` en un editor de texto
2. Busca la línea:
   ```javascript
   const API_URL = 'PEGA_AQUÍ_TU_URL_DEL_WEB_APP';
   ```
3. Reemplaza con tu URL real del paso 2
4. Repite lo mismo en `admin.html`

### Paso 4 — Configurar el token de admin

1. En tu Google Sheets, ve a la hoja `configuracion`
2. Busca la fila `admin_token`
3. Cambia el valor por algo seguro (ej: `MiT0k3nS3gur0!2024`)
4. Este token es tu contraseña para entrar al panel admin

### Paso 5 — Probar localmente

1. Abre `tienda.html` directamente en tu navegador (doble click)
2. Verifica que carga los productos de ejemplo
3. Abre `admin.html` e ingresa con tu token
4. Crea un producto de prueba y verifica que aparece en la tienda

---

## 🌐 Publicar en GitHub Pages (gratis)

1. Crea una cuenta en [github.com](https://github.com) si no tienes
2. Crea un repositorio nuevo llamado `mi-tienda` (público)
3. Sube estos archivos: `index.html`, `tienda.html`, `admin.html`, `404.html`
4. Ve a **Settings > Pages** del repositorio
5. En **Source**, selecciona **Deploy from a branch**
6. Selecciona **main** y **/ (root)**
7. Espera ~2 minutos y tu tienda estará en:
   ```
   https://TU-USUARIO.github.io/mi-tienda/
   ```

---

## 🗺️ Google Maps (opcional)

Para activar el autocompletado de dirección:

1. Ve a [Google Cloud Console](https://console.cloud.google.com)
2. Crea un proyecto nuevo (o usa uno existente)
3. Ve a **APIs & Services > Library**
4. Busca y habilita **Places API**
5. Ve a **APIs & Services > Credentials**
6. Click en **Create Credentials > API Key**
7. Copia la API Key
8. En `tienda.html`, busca la línea al final:
   ```html
   <script src="https://maps.googleapis.com/maps/api/js?key=TU_API_KEY_AQUI...
   ```
9. Reemplaza `TU_API_KEY_AQUI` con tu API Key real

> **Nota**: Sin la API Key de Maps, el campo de dirección funciona como texto normal. La tienda funciona perfectamente sin esta característica.

---

## 🔒 Notas de seguridad

- **No compartas tu `admin_token`** con nadie
- El token se almacena en `sessionStorage` del navegador (se borra al cerrar la pestaña)
- Las API Keys de Google Maps deben restringirse a tu dominio en Google Cloud Console
- El `Code.gs` nunca expone el `admin_token` ni `admin_email` públicamente

---

## 📋 CORS en Apps Script

Google Apps Script maneja CORS de una manera especial. Si tienes problemas de CORS:

1. Al desplegar como Web App, asegúrate de seleccionar **"Cualquier persona"** en acceso
2. Los requests se redirigen automáticamente, por lo que `fetch` con `mode: 'cors'` funciona
3. Si hay problemas, usa `mode: 'no-cors'` (pero perderás acceso al body de la respuesta)
4. La solución más robusta es usar `fetch(url, { redirect: 'follow' })`

> **Importante**: Cada vez que modifiques el `Code.gs`, debes crear una **nueva implementación** (no actualizar la existente) para que los cambios se reflejen. Ve a **Implementar > Administrar implementaciones > Nueva implementación**.

---

## 💡 Tips

- **Imágenes**: Puedes usar Google Drive para alojar imágenes. Sube la imagen, hazla pública, y usa la URL: `https://drive.google.com/uc?export=view&id=ID_DEL_ARCHIVO`
- **Actualizar la tienda**: Cada vez que cambies datos en el admin, los cambios se reflejan inmediatamente en la tienda al recargar
- **Backup**: Tu base de datos es el Google Sheets, puedes descargarlo como Excel en cualquier momento
- **Notificaciones**: Configura `admin_email` en la hoja de configuración para recibir un email con cada nuevo pedido

---

## 🐛 Solución de problemas

| Problema | Solución |
|----------|----------|
| "Error al cargar productos" | Verifica que la URL del Web App sea correcta y esté desplegada |
| "Token incorrecto" en admin | Verifica el valor de `admin_token` en la hoja `configuracion` |
| Los cambios en Code.gs no se reflejan | Crea una NUEVA implementación, no actualices la existente |
| CORS error | Asegúrate de que el Web App tiene acceso "Cualquier persona" |
| WhatsApp no abre | Verifica que `telefono_whatsapp` tenga formato `57XXXXXXXXXX` |

---

Hecho con ❤️ usando Google Apps Script + HTML/CSS/JS Vanilla
