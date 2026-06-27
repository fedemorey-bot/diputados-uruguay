# Metodología del proyecto — Seguimiento parlamentario Uruguay

Documento de referencia para no tener que re-descubrir todo cada vez.
Caso piloto: Conrado Rodríguez (Partido Colorado, Montevideo, ID 11597).

---

## 1. Arquitectura general

- **Sitio**: estático (HTML/CSS/JS), sin backend.
- **Hosting**: GitHub Pages, repo `github.com/fedemorey-bot/diputados-uruguay`.
- **Despliegue**: automático vía GitHub Actions (`.github/workflows/static.yml`)
  cada vez que se sube un archivo nuevo a la rama `main`.
- **Archivo principal**: `index.html` (todo el sitio en un solo archivo,
  datos embebidos como JS).

---

## 2. Fuentes de datos y scripts

### 2.1. Intervenciones en el PLENARIO
- **Script**: `scraper_pdf_2025_2026.py`
- **Fuente**: `documentos.diputados.gub.uy/docs/L{legislatura}/Diario/d{N}.pdf`
  - Numeración predecible y continua. Legislatura 50 arranca en N=4548 (15/02/2025).
  - El script auto-detecta dónde quedó la última vez (lee la planilla existente)
    y se detiene solo cuando encuentra ~8 números seguidos que no existen
    (asume que llegó al presente). **Por eso se puede correr de nuevo sin tocar nada.**
  - Algunas sesiones largas se publican en varios "Tomos" separados → el script
    avisa cuáles quedaron pendientes, no es un error.
- **Patrón de orador**: `SEÑOR|SEÑORA APELLIDO.-` (a veces con alias entre
  paréntesis: `SEÑOR RODRÍGUEZ (don Edgardo).-` — hay que permitir ese paréntesis
  opcional en el regex o se pierde el match).
- **Quién preside (importante)**: el diario marca cada cambio así:
  `(Ocupa la Presidencia el señor/la señora representante NOMBRE.)`
  Hay que recorrer el texto en orden e ir actualizando quién preside en cada
  momento (no alcanza con un solo nombre fijo para todo el diario). Si no hubo
  ningún relevo, se usa como respaldo la firma de cierre del diario
  (`Se levanta la sesión... NOMBRE PRESIDENTE/PRESIDENTA`).
- **Fecha**: patrón `Montevideo, DD de MES de AAAA`.
- **Corte de texto**: cortar todo lo que viene después de
  "Se levanta la sesión" (después de eso solo hay votaciones electrónicas y
  anexos con proyectos completos — si no se corta, se mezcla todo).
- **Caracteres ilegales**: los PDF a veces traen caracteres de control que
  rompen Excel — hay que limpiarlos antes de guardar (`\x00-\x08`, etc.).
- **Salida**: `intervenciones_diputados_uruguay.xlsx`

### 2.2. Intervenciones en COMISIONES (Versiones Taquigráficas)
- **Script**: `scraper_comisiones.py`
- **Fuente**: `parlamento.gub.uy/camarasycomisiones/{ambito}/comisiones/{id}/versiones-taquigraficas`
  - `{ambito}` es **casi siempre `representantes`**, PERO algunas comisiones
    especiales (ej. la Carcelaria) dependen de la **Asamblea General**, no de
    la Cámara → usan `asambleageneral` en su lugar. Hay que revisar la ficha
    de cada comisión para saber cuál le corresponde (el usuario puede pasar
    el link de la comisión y ahí se ve en la URL).
  - Cada documento real está en:
    `documentosyleyes/documentos/versiones-taquigraficas/{ambito}/50/{numero}/0`
    (página intermedia que a su vez tiene el PDF real en
    `infolegislativa.parlamento.gub.uy/temporales/{numero}.PDF` — URL fija y
    predecible, a diferencia de los diarios de sesión).
- **Patrón de orador**: igual que el plenario, PERO también aparece con
  prefijos de rol que hay que sacar: `SEÑOR REPRESENTANTE APELLIDO (Nombre).-`
  y `SEÑOR SENADOR APELLIDO (Nombre).-` (si no se sacan, el nombre queda
  mezclado con la palabra "REPRESENTANTE"/"SENADOR" y no matchea con nadie).
- **Fecha (distinta a la del plenario)**: a veces sí dice "Montevideo, ..."
  pero el formato típico es:
  `Versión taquigráfica de la reunión realizada el día [día] DD de MES de AAAA`
  → hace falta un segundo patrón de fecha además del de plenario.
- **Cierre**: puede decir "Se levanta la sesión" O "Se levanta la reunión".
- **Comisiones de Conrado Rodríguez (8)**: Hacienda (id 78), Presupuestos
  integrada con Hacienda (1310), Hacienda+Transporte (1324), Lavado de
  Activos (1271), Población y Desarrollo (1270), Educación y Cultura (211),
  Legislación del Trabajo (1171), Carcelaria (1301, ámbito asambleageneral).
- **Salida**: `intervenciones_comisiones.xlsx` (separada de la del plenario)
- **Pendiente conocido**: ~50% de las filas todavía no tienen fecha — falta
  identificar un tercer formato de fecha que usan algunos documentos viejos.

### 2.3. Roster de legisladores (para identificar nombres)
- **Titulares (99)**: `parlamento.gub.uy/camarasycomisiones/representantes/plenario/integracion/titulares`
  (nombre, ID, partido, departamento, foto).
- **Suplentes activos actualmente**: misma sección sin `/titulares`
  (`.../plenario/integracion`) — solo trae a los que están sustituyendo HOY,
  no el historial completo. El historial completo
  (`/sobreelparlamento/integracionhistorica`) a veces bloquea el acceso
  automatizado (bot detection) — no se pudo resolver del todo.
- **Cómo identificar un nombre crudo** (con posibles espacios raros de la
  extracción del PDF, o nombre incompleto): normalizar (sin tildes, sin
  espacios, mayúsculas) y probar: (1) match exacto, (2) si no, buscar por
  "contención" (el nombre buscado está contenido en el nombre completo del
  roster, o viceversa) — esto resuelve apellidos compuestos y nombres
  parciales. Si hay 2+ candidatos, NO adivinar, dejar el texto crudo.

### 2.4. Ficha individual del legislador ("Despacho Virtual")
- **URL base**: `parlamento.gub.uy/camarasycomisiones/legisladores/{id}`
- **Sub-páginas con descarga JSON** (agregar `/json?Fecha_desde=X&Fecha_hasta=Y&page&_format=json`):
  - `/actuacion-legislador` — log completo de actividad (votos, designaciones,
    intervenciones, con referencia a tomo/página/diario).
  - `/iniciativas-legislador` — proyectos de ley presentados.
  - `/pedidosInf-legislador` — pedidos de informes (con asunto N°, organismo,
    fecha de respuesta).
  - `/asistenciaplenario` — asistencia a sesiones de plenario.
  - `/asistencia-a-comisiones` — asistencia por comisión.
- **¡OJO con el caché!**: `web_fetch` a veces devuelve una versión vieja
  cacheada de estas páginas (sobre todo si se pide un rango de fechas
  distinto al de una búsqueda anterior). Si los datos no coinciden con lo
  esperado, usar el navegador (Claude in Chrome) en vez de `web_fetch` —
  el navegador sí carga la página real y actualizada.

### 2.5. Videos de YouTube de sesiones del PLENARIO
- **Índice oficial**: `parlamento.gub.uy/camarasycomisiones/representantes/plenario/documentos/diarios-de-sesion/videos`
  (paginado con `?page=N`, ~10-12 items por página, mezclado con contenido
  que no es sesión).
- **Las comisiones NUNCA se transmiten** — solo el plenario tiene video.
- **BUG GRAVE que ya resolvimos**: emparejar video↔fecha con un regex que
  busca "el primer título después del link" **falla** (puede asignar el
  video equivocado a una fecha). La forma correcta es usar la estructura real
  del HTML con BeautifulSoup:
  - El link está en: `.views-field-field-ytp-thumbnail-high a[href*="watch?v="]`
  - El título está en: `.views-field-title .field-content`
  - Ambos viven DENTRO del mismo contenedor de tarjeta — hay que subir desde
    el link hasta encontrar el título correspondiente, no buscar a ciegas en
    todo el texto de la página.
  - Script ya corregido: `indexar_videos_sesiones.py`
- **Minuto exacto dentro del video**: algunos videos (no todos) tienen
  capítulos puestos a mano en la descripción, con el nombre del orador y su
  minuto (ej: "10:06:29 Rodríguez, Conrado") — hay que abrir el video,
  expandir la descripción ("...más") y buscar el nombre ahí.
  - Intentamos automatizar esto cruzando la transcripción automática de
    YouTube (vía librería `youtube-transcript-api`) con el texto real de la
    intervención (`buscar_minuto_exacto.py`), pero la tasa de éxito fue baja
    (problemas de transcripción automática / reconocimiento de voz). **Queda
    pendiente, no resuelto del todo.**

### 2.6. Viajes al exterior
- No hay una fuente oficial única y estructurada por legislador.
- El "Informe de Viajes" oficial (`diputados.gub.uy/transparencia/informe-de-viajes/`)
  cubre sobre todo viajes a PARLATINO/Parlasur, no visitas bilaterales.
- Mejor método encontrado: búsqueda de prensa por nombre + "viaje"/"misión".

### 2.7. Redes sociales
- Instagram, Facebook, Wikipedia: alcanza con un link simple, sin problema.
- **X/Twitter — NO HAY forma gratuita confiable de mostrar posts en vivo.**
  El widget oficial gratuito de embed está roto/poco confiable desde 2023, y
  confirmado que sigue así en 2026 (no es un error nuestro, es un problema de
  la plataforma). Decisión tomada: dejar solo un link a su perfil, sin
  intentar embeber contenido.

---

## 3. Lecciones operativas (flujo de trabajo en la Mac del usuario)

- El navegador **renombra automáticamente** los archivos duplicados en
  Descargas (`archivo.py`, `archivo_1.py`, `archivo_2.py`...). Esto causó
  MUCHA confusión — varias veces se corrió una versión vieja sin querer.
  **Siempre verificar con `grep` o por tamaño de archivo antes de correr.**
- Cada vez que se abre una ventana **nueva** de Terminal, hay que volver a
  hacer `cd ~/Desktop/Diputados` (no lo recuerda de la sesión anterior).
- Los scripts de diarios de sesión (plenario) tienen auto-resume; los de
  comisiones y videos NO — hay que borrar el Excel antes de re-correr o se
  duplican los datos.
- GitHub Pages tarda 1-3 minutos en publicar después de subir un archivo.
  `raw.githubusercontent.com` tiene su PROPIO caché de CDN, independiente —
  puede mostrar contenido viejo aunque el commit ya esté actualizado.
- Subir archivos por la web de GitHub ("Add file → Upload files") puede
  introducir errores de tipeo en el nombre del archivo (pasó con
  `index..html`, dos puntos en vez de uno) — siempre confirmar el nombre
  exacto después de subir.

---

## 4. Pendientes conocidos (al cierre de esta etapa)

1. Completar fechas faltantes en intervenciones de comisiones (~50%)
2. Integrar las intervenciones de comisiones a la página web de Conrado
3. Terminar (o aceptar las limitaciones de) el buscador de minuto exacto
4. Verificar cuántas de las intervenciones de comisiones son específicamente
   de Conrado (la planilla trae a TODOS los que hablaron, no solo a él)
5. Replicar todo este proceso para el resto de los 99 diputados — **a
   propósito, pausado hasta terminar de pulir este caso piloto**
