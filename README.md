# App de Finanzas (PWA)

## Publicar en GitHub Pages
1. Sube estos 5 archivos a la raíz de tu repositorio (index.html, sw.js, manifest.webmanifest, icon.png, mod2.js), reemplazando los que ya tenías.
2. Espera 1-2 minutos a que GitHub Pages termine de publicar.
3. Abre tu enlace de GitHub Pages en Safari.
4. Si ya tenías la app instalada, ábrela: te preguntará "Hay una versión nueva ¿Actualizar ahora?" — toca Aceptar. Si no aparece, cierra la app desde el multitarea y ábrela de nuevo.

## Cómo recibir cada cambio nuevo
Cada vez que pidas un cambio aquí, te entregaré de nuevo este mismo paquete completo, ya con el número de versión del service worker (V='finanzas-vN') incrementado en sw.js. Reemplaza los archivos del repositorio por los nuevos (mismo nombre, se sobrescriben) y vuelve a abrir la app: el aviso de actualización aparece solo. No necesitas tocar nada de configuración de GitHub cada vez, ni reinstalar el ícono.

## No se sube al repositorio
`datos-iniciales.json` u otras copias de seguridad con tus datos reales: el repositorio es público y ese archivo solo se importa desde dentro de la app, nunca se publica en Pages.
