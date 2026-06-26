# Sistema de Evaluación de Desempeño — TOOLTEK® SpA

Sistema simultáneo de autoevaluación (trabajador) + evaluación de jefatura, con informe unificado, previsualización editable para el evaluador y envío automático por correo.

## Archivos

- `evaluacion_desempeno.html` → la app (subir a GitHub Pages)
- `Code.gs` → backend (pegar en Google Apps Script)
- `Plantilla_Evaluacion_Desempeno.xlsx` → estructura base para tu Google Sheet

## Cómo se calcula la nota (idéntico a tu Excel actual)

```
CDC        = Comprender×20% + Desear×20% + Capacidad×20%
Valores    = Trabajamos×10% + Excelencia×10% + Desafío×10% + Confianza×10%

Resultado Autoevaluación (ya ponderado) = (CDC_auto + Valores_auto) × 40%
Resultado Jefatura       (ya ponderado) = (CDC_jefe + Valores_jefe) × 60%

NOTA FINAL = Resultado Autoevaluación + Resultado Jefatura
```

Escala: 0–33,4 Mejora Necesaria · 33,5–66,6 Eficaz · 66,7–100 Cumple/Supera Expectativas.

Validado contra tu PDF de ejemplo (Francisco Puga): da exactamente **57,8**.

---

## Paso 1 — Crear el Google Sheet

1. Sube `Plantilla_Evaluacion_Desempeno.xlsx` a Google Drive.
2. Ábrelo y guárdalo como Google Sheets (Archivo → Guardar como Hojas de cálculo de Google).
3. Verás 3 pestañas:
   - **Nomina**: completa aquí cada trabajador con su jefatura asignada (RUT, nombre, cargo, email, periodo). Ya tiene 2 filas de ejemplo — bórralas cuando cargues la real.
   - **Config**: pon el correo real de RR.HH. en `Email_RRHH`.
   - **Procesos**: NO la toques. Se llena sola cuando la gente empieza a evaluar.

> Nota: puedes correr el mismo proceso para varios periodos. Solo agrega filas nuevas en "Nomina" con el `Periodo` actualizado (ej. "3er Trimestre 2026"); el sistema crea automáticamente un nuevo proceso por cada combinación Evaluado+Periodo.

## Paso 2 — Desplegar el backend (Apps Script)

1. En tu Google Sheet: **Extensiones → Apps Script**.
2. Borra el contenido por defecto y pega todo el contenido de `Code.gs`.
3. Guarda (ícono disco).
4. Haz clic en **Implementar → Nueva implementación**.
5. Tipo: **Aplicación web**.
6. Configuración:
   - Ejecutar como: **Yo (tu cuenta)**
   - Quién tiene acceso: **Cualquier usuario** (para que evaluado/evaluador externos al dominio puedan entrar; si todos tienen Google Workspace de TOOLTEK, puedes restringir a "Cualquier usuario de tu organización")
7. Autoriza los permisos cuando te lo pida (acceso a Sheets + envío de correo con tu cuenta).
8. Copia la **URL del Web App** que te entrega (termina en `/exec`).

## Paso 3 — Conectar el HTML con el backend

1. Abre `evaluacion_desempeno.html` con un editor de texto.
2. Busca la línea:
   ```js
   const API_URL = 'PEGAR_AQUI_URL_DEL_WEB_APP_DE_APPS_SCRIPT';
   ```
3. Reemplaza por la URL que copiaste en el Paso 2.8.
4. Guarda.

## Paso 4 — Publicar en GitHub Pages

Mismo flujo que ya usas en tus otros proyectos (Cargas_Transito, Inventario_Ropa):

1. Crea/usa un repo en `tooltekspa.github.io` (o sube el archivo a un repo existente).
2. Sube `evaluacion_desempeno.html`.
3. Activa GitHub Pages (legacy build) si no está activo.
4. La URL final quedará algo como:
   `https://tooltekspa.github.io/Evaluacion_Desempeno/evaluacion_desempeno.html`

---

## Cómo funciona el flujo para el usuario final

1. Evaluado y evaluador entran al **mismo link**, cada uno ingresa su RUT.
2. El sistema detecta automáticamente su rol (o les deja elegir si tienen varios procesos pendientes).
3. Ambos pueden completar su formulario **al mismo tiempo**, sin pisarse — cada uno escribe en su propia "mitad" en la hoja Procesos.
4. Cuando **ambos** terminaron, el sistema calcula la nota final automáticamente.
5. El **evaluador (jefatura)** puede entonces ver el botón "Ver previsualización del informe":
   - Ve el informe completo armado (igual al formato que ya usas).
   - Puede editar su comentario final antes de enviar.
   - Solo él tiene el botón "Confirmar y enviar informe por correo".
6. Al confirmar, el sistema envía un correo HTML con el informe a: **evaluado + evaluador + RR.HH.** (definido en la pestaña Config).
7. El evaluado, si entra antes de que el evaluador confirme, solo puede **ver** la previsualización (modo lectura), no enviarla.

## Qué puedes personalizar después

- **Pesos de competencias**: están hardcodeados en `Code.gs` función `calcularBloque()` (20%/20%/20% y 10%×4). Si cambian, edítalos ahí.
- **Escala de colores**: función `clasificar()` en `Code.gs`.
- **Destinatarios del correo**: función `confirmarEnvio()` — actualmente evaluado + evaluador + RR.HH. Si quieres agregar al Gerente de Administración y Finanzas siempre en copia, agrega su correo fijo ahí.
- **Múltiples periodos en paralelo**: ya soportado vía la columna `Periodo` en Nomina.

## Limitación a tener en cuenta

Si el mismo RUT aparece como evaluador de **varias** personas en el mismo periodo, al loguearse verá una lista para elegir a qué trabajador evaluar (pantalla "Tienes más de un proceso disponible"). Esto es intencional para que jefes con varios trabajadores a cargo puedan navegar entre evaluaciones.
