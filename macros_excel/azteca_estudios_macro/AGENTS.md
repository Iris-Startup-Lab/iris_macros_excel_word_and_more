# AGENTS.md — Guía para Agentes de IA en Azteca Estudios Macro

Este documento proporciona contexto técnico, arquitectura, convenciones y reglas operativas para agentes de IA (y desarrolladores) que interactúen, mantengan o extiendan este repositorio.

---

## 1. Visión General del Proyecto

**Nombre:** Excel VBA Automation Toolkit — Cotizaciones Azteca Estudios  
**Propósito:** Sistema de automatización en VBA para Microsoft Excel en Windows orientado a la generación, estructuración, cálculo, colapso visual y exportación a PDF de cotizaciones comerciales de Azteca Estudios.

### Funcionalidades Clave

1. **UUID Estable y Automático:** Generación de un UUID único por cotización en la celda `K4`, vinculado al nombre del proyecto (`C4`) y persistido en una hoja de sistema muy oculta (`Meta`).
2. **Esquema Multinivel (Agrupación/Colapso):** Colapso y expansión a demanda de categorías completas (Nivel 1) y líneas de detalle técnico (Nivel 2, identificadas por la marca `"D"` en la columna oculta `O`).
3. **Exportación Dual a PDF en 1 Clic:** Generación automática de dos archivos PDF (`..._Resumen.pdf` y `..._Detalle.pdf`) con nomenclatura estandarizada.
4. **Cálculo Dinámico de Cotizaciones:** Recálculo automático de subtotales por bloque de categoría, Subtotal General, IVA (16%) y Total Final (`suma.vb`).
5. **Manipulación Dinámica de Estructura:** Inserción y eliminación de filas, adición de fechas/ítems y duplicación de bloques de categorías desde la hoja `Plantillas`.
6. **Blindaje Operativo (`M_Blindaje`):** Anclaje, registro y bloqueo de los botones de la hoja; reimpresión automática del ID Cotización si se borra; y protección de la estructura del libro para impedir la creación o copia de hojas de cotización.

---

## 2. Mapa y Estructura del Repositorio

```
azteca_estudios_macro/
├── AGENTS.md                                # Este documento de contexto para agentes de IA
├── README.md                                # Documentación técnica general y visión del proyecto
├── instrucciones.md                         # Guía paso a paso de instalación y solución de problemas para usuarios
├── GUÍA DE USO COTIZACIÓN.pdf              # Manual visual de usuario en PDF
│
├── Plantilla Cotización AE 1.xlsm           # Plantilla principal de trabajo (habilitada para macros)
├── Plantilla Cotización AE 1_para_demo.xlsm  # Versión de demostración
├── Plantilla Cotización AE 1_bkp.xlsm       # Backup de la plantilla
├── Macro_generadora_uuids.xlsm              # Libro auxiliar con prueba de UUID
│
├── bas/                                     # Exportaciones .bas para importar directamente al editor de VBA (Alt+F11)
│   ├── macro_generadora_uuid.bas
│   ├── asignar_uuid_cotizacion.bas
│   ├── esquema_categorias.bas
│   ├── generar_pdfs_resumen_detalle.bas
│   └── blindaje_botones.bas
│
├── Módulos Principales de la Nueva Versión (.vb):
│   ├── macro_generadora_uuid.vb            # Función base GENERAR_UUID() mediante API de Windows (ole32.dll)
│   ├── asignar_uuid_cotizacion.vb          # Módulo M_UUID (Gestión de K4, C4, hoja Meta y constantes globales)
│   ├── esquema_categorias.vb               # Módulo M_Esquema (Agrupación Outline nivel 1 y 2, columna O)
│   ├── generar_pdfs_resumen_detalle.vb     # Módulo M_PDF (Exportación de _Resumen.pdf y _Detalle.pdf)
│   ├── blindaje_botones.vb                 # Módulo M_Blindaje (Anclaje/registro/bloqueo de botones, protección de estructura)
│   └── eventos_uuid.vb                     # Código de eventos (ThisWorkbook y hoja Cotización)
│
└── Macros Utilitarias Secundarias/Legacy (.vb):
    ├── suma.vb                             # Recálculo de totales por categoría, Subtotal, IVA y Total
    ├── duplicar_bloque.vb                  # Copia estructuras de bloque desde la hoja "Plantillas"
    ├── agregar_fila_item.vb                # Inserta fila simple dentro de una categoría
    ├── agregar_fila_producto.vb            # Inserta fila de producto formateada
    ├── agregar_fila_fecha.vb               # Inserta fila con campo fecha
    ├── borrar_fila_seleccionada.vb         # Elimina fila activa manteniendo consistencia
    ├── limpiar_cotizacion.vb               # Resetea campos de la cotización
    ├── generar_pdf.vb                      # Exportación simple a PDF
    └── generar_pdf_nombre_inteligente.vb   # Exportación PDF con nombre derivado de celdas
```

---

## 3. Arquitectura y Componentes Clave

### A. Módulo `M_UUID` (`asignar_uuid_cotizacion.vb` / `.bas`)

- **Celda Destino:** `K4` ("ID Cotización"). Formateada explícitamente como texto (`@`).
- **Celda Disparadora:** `C4` ("Nombre del Proyecto").
- **Celda Versión:** `N4`.
- **Hoja Meta (`Meta`):** Creada automáticamente si no existe (vía `HojaMeta()`, ahora **pública**). Configurada como `xlSheetVeryHidden` (invisible en la interfaz de Excel, no desocultable manualmente por el usuario). Guarda `(Proyecto, UUID, Fecha)` en `B1:B3` y el registro de posiciones de botones en `D:H`. Su creación levanta `gCreandoHojaSistema` y desprotege/reprotege la estructura del libro.
- **`Meta!B2` es la fuente de la verdad; `K4` es solo su proyección.** Corolario obligatorio: nunca generes un UUID nuevo para "arreglar" un `K4` vacío o alterado — reimprime el de `Meta`. `AsegurarUUIDCotizacion` solo invoca `GENERAR_UUID()` si `Meta!B2` está vacío o si `C4` cambió.
- **Vigilancia del ID (`VigilarIDCotizacion(Target)`):** Invocado por `Worksheet_Change`. Intersecta `Target` contra `C4,K4`. Reentrante-seguro mediante `Static ocupado`; restaura `Application.EnableEvents = True` en todas las rutas de salida (incluida la etiqueta `salir:`). Notifica al usuario solo cuando se tocó `K4` y el valor cambió.
- **Red de seguridad (`RestaurarIDSiFalta`):** Versión silenciosa, llamada desde `Workbook_Open`, `BeforeSave`, `BeforePrint` y `GenerarPDFsCotizacion`.
- **`UUIDGuardado()`:** Lee `Meta!B2` sin generar nada. Úsala para comparaciones; nunca `AsegurarUUIDCotizacion` con ese fin.
- **Constantes Compartidas:**
  - `COL_MARCA = "O"`: Columna oculta que contiene la marca de detalle.
  - `MARCA_DETALLE = "D"`: Valor que identifica una fila como detalle técnico colapsable.
  - `FILA_INICIO = 7`: Primera fila a evaluar en el cuerpo de la cotización.
- **Helper de Búsqueda de Hojas (`HojaQueContiene` / `HojaCotizacion`):** Búsqueda flexible e insensible a mayúsculas/acentos/espacios para prevenir fallos por renombrado parcial de hojas (ej. `"Cotización "`).

### B. Módulo `M_Esquema` (`esquema_categorias.vb` / `.bas`)

- **Nivel 1 (Categoría):** Agrupa desde la fila posterior al título de la categoría hasta la fila previa a `SUMA TOTAL`. Permite colapsar el bloque completo dejando visible solo el encabezado y su total.
- **Nivel 2 (Detalle):** Detecta rachas consecutivas de filas con valor `"D"` en la columna `O` y aplica `.Group`.
- **Carga de Categorías (`CargarCategorias`):** Lee la hoja `Categorías` para identificar dinámicamente los nombres válidos de categoría.
- **MergeArea Handling:** `FilaFinEncabezado` determina correctamente el final de títulos fusionados (ej. `A7:N8`).

### C. Módulo `M_PDF` (`generar_pdfs_resumen_detalle.vb` / `.bas`)

- Exporta dos documentos PDF en la misma carpeta del archivo `.xlsm`:
  1. `..._Detalle.pdf`: Muestra todas las filas visibles y detalladas.
  2. `..._Resumen.pdf`: Oculta temporalmente las filas con marca `"D"` antes de exportar y restaura su visibilidad inmediatamente después.
- **Formato del Nombre:** `Cotizacion_<Cliente>_<Proyecto>_<UUID>_v<Version>_<Fecha>_<Tipo>.pdf`. Los caracteres no válidos para rutas Windows (`\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`) son higienizados automáticamente.

### D. Módulo `M_Blindaje` (`blindaje_botones.vb` / `.bas`)

Resuelve tres fallas operativas. **Léelo antes de tocar cualquier macro que inserte, borre u oculte filas.**

- **Problema 1 — botones que se desplazan al área imprimible.** Causa raíz: los objetos nacen con `Placement = xlMoveAndSize`, es decir anclados a las celdas subyacentes. Todo `Rows(x).Insert`, `Rows(x).Delete`, `EntireRow.Hidden = True` y `.Group` los arrastra, aplasta o —cuando la inserción proviene de `Rows(x).Copy`— los **duplica**.
  - `FijarBotones(ws)`: fuerza `Placement = xlFreeFloating` y `Locked = True`. Idempotente; sale sin escribir si ya está todo anclado (para no marcar el libro como modificado).
  - `GuardarPosicionesBotones()`: persiste `nombre / Left / Top / Width / Height` de cada `Shape` en `Meta!D:H`.
  - `RestaurarPosicionesBotones([silencioso])`: reposiciona con tolerancia de `0.5` pt; si nada se movió no escribe nada.
  - `LimpiarBotonesDuplicados()`: elimina los `Shape` ausentes del registro, previa confirmación.
  - `BloquearBotones` / `DesbloquearBotones`: `ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=False`. **`Contents:=False` es deliberado:** protege únicamente los objetos y deja las celdas editables, por lo que ninguna macro existente requiere desproteger/reproteger.
- **Problema 2 — borrado del ID Cotización.** Ver `M_UUID.VigilarIDCotizacion` y `M_UUID.RestaurarIDSiFalta`.
- **Problema 3 — creación de hojas de cotización.** `ProtegerEstructura` / `DesprotegerEstructura` aplican `ThisWorkbook.Protect Structure:=True`, complementado por el evento `Workbook_NewSheet`.
- **Constantes y estado global:**
  - `PWD_BLINDAJE`: contraseña única para protección de hoja y de estructura. Si un agente añade otra protección, **debe** reutilizar esta constante.
  - `gCreandoHojaSistema`: bandera pública que `HojaMeta()` levanta para que `Workbook_NewSheet` no elimine la hoja de sistema `Meta`. Toda macro que necesite crear una hoja legítima debe levantarla y bajarla.
- **Entradas:** `BlindarLibro()` (puesta en marcha) y `DesblindarLibro()` (mantenimiento).

### E. Eventos (`eventos_uuid.vb`)

- **`ThisWorkbook`:**
  - `Workbook_Open`: `AsegurarUUIDCotizacion` → `FijarBotones` → `RestaurarPosicionesBotones` → `ProtegerEstructura`.
  - `Workbook_NewSheet(Sh)`: si `gCreandoHojaSistema` es `False`, elimina `Sh` (con `DisplayAlerts`/`EnableEvents` desactivados), reprotege la estructura y notifica que una cotización equivale a un archivo `.xlsm`.
  - `Workbook_BeforeSave` y `Workbook_BeforePrint`: `RestaurarIDSiFalta` + `RestaurarPosicionesBotones` (ningún archivo se guarda ni imprime sin ID ni con botones fuera de lugar).
- **Hoja `Cotización`:**
  - `Worksheet_Change`: delega en `VigilarIDCotizacion Target`, que vigila `C4` **y `K4`**.
  - `Worksheet_Activate`: `FijarBotones Me` + `RestaurarPosicionesBotones`.

> Los eventos son deliberadamente delgados (una línea de delegación): toda la lógica vive en módulos estándar para que sea importable, versionable y testeable.

### F. Recálculo Matemático (`suma.vb`)

- Recorre la columna `A` buscando etiquetas `"SUMA TOTAL"` o `"SUMA TOTAL:"`.
- Acumula los valores numéricos de la columna `M` (Total por renglón) hacia la celda `M` de la fila de suma.
- Busca en la columna `I` los rótulos `"SUBTOTAL"`, `"IVA"` y `"TOTAL"` para asignar en la columna `K`:
  - `SubtotalGeneral` = $\sum \text{Totales Categorías}$
  - `IVA` = $\text{SubtotalGeneral} \times 0.16$
  - `TotalFinal` = $\text{SubtotalGeneral} + \text{IVA}$

---

## 4. Convenciones Técnicas y Reglas para Agentes

Al modificar o crear código en este repositorio, **sigue estrictamente las siguientes reglas**:

### 1. Compatibilidad de Entorno

- **Sistema Operativo Target:** Windows. La generación de UUIDs depende de las funciones Windows API (`CoCreateGuid` y `StringFromGUID2` en `ole32.dll`).
- **Formato de Libro:** Todo archivo funcional debe ser `.xlsm` (Libro de Excel Habilitado para Macros). Nunca guardar versiones activas como `.xlsx`.

### 2. Sincronización Doble de Código (`.vb` y `bas/*.bas`)

- Los archivos `.vb` en la raíz son la fuente principal de código legible en UTF-8 / Markdown.
- Los archivos `.bas` en la carpeta `bas/` corresponden a los módulos estándar listos para importación en el editor VBA de Excel. **Formato obligatorio:** codificación **cp1252 (Windows-1252)**, saltos de línea **CRLF**, y primera línea `Attribute VB_Name = "<NombreDelMódulo>"`. Si se escriben en UTF-8, los acentos aparecen corruptos dentro del editor VBA.
- **REGLA:** Si realizas un cambio en una macro modular (`asignar_uuid_cotizacion`, `esquema_categorias`, `generar_pdfs_resumen_detalle`, `macro_generadora_uuid`, `blindaje_botones`), **debes actualizar tanto el archivo `.vb` en la raíz como el correspondiente `.bas` dentro de `bas/`**.
- Mapeo archivo → nombre de módulo: `asignar_uuid_cotizacion` → `M_UUID`, `esquema_categorias` → `M_Esquema`, `generar_pdfs_resumen_detalle` → `M_PDF`, `blindaje_botones` → `M_Blindaje`.

### 2.bis Regla de Integridad Visual (obligatoria)

Toda macro que **inserte, elimine, oculte, agrupe o copie filas** en la hoja `Cotización` **debe**:

1. Llamar `FijarBotones ws` (o `FijarBotones` sin argumentos) **antes** de la operación, para que los objetos no se arrastren, aplasten ni dupliquen.
2. Llamar `RestaurarPosicionesBotones` **al final**, como red de seguridad.

Macros que ya cumplen: `AgregarFilaProducto`, `AgregarFilaFecha`, `AgregarFilaItem`, `BorrarFilaSeleccionada`, `DuplicarBloque`, `ConstruirEsquema`, `ResetMacroDefault`, `GenerarPDFsCotizacion`. **Si añades una nueva, replica el patrón.**

Ambas rutinas son idempotentes y no marcan el libro como modificado si no hay nada que cambiar, así que llamarlas de más es inofensivo.

### 3. Buenas Prácticas de Código VBA

- **`Option Explicit` OBLIGATORIO:** Todas las variables deben declararse explícitamente (`Dim`).
- **Manejo de Pantalla y Eventos:** Al realizar operaciones masivas o de ocultación temporal (como en esquemas o generación de PDFs):

  ```vba
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  ' ... código ...
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  ```

- **Control de Errores Limpio:** Usar `On Error Resume Next` / `On Error GoTo 0` o etiquetas de salida limpias (`On Error GoTo salir`) garantizando que `EnableEvents` y `ScreenUpdating` se restablezcan en `True` siempre.
- **Acceso Robusto a Hojas:** No asumas el nombre exacto de la pestaña ni el CodeName interno (`Hoja1`). Usa la función helper `HojaCotizacion()` / `HojaQueContiene("COTIZACI")` para tolerar variaciones de espacios o acentos.

### 4. Mapeo Rígido de Celdas y Columnas

| Elemento / Campo | Ubicación en Excel |
| --- | --- |
| Nombre del Cliente | Celda `C3` |
| Nombre del Proyecto | Celda `C4` (Dispara UUID) |
| ID Cotización (UUID) | Celda `K4` (Formato texto `@`) |
| Versión | Celda `N4` |
| Columna de Estructura / Categorías | Columna `A` |
| Rótulos Totales Generales | Columna `I` (`SUBTOTAL`, `IVA`, `TOTAL`) |
| Subtotales de Renglón / Categoría | Columna `M` |
| Valores Totales Generales | Columna `K` |
| Marca Oculta de Detalle | Columna `O` (`"D"` = Fila de detalle) |
| Hoja Meta de Sistema | `Worksheets("Meta")` (`xlSheetVeryHidden`) |
| Proyecto atado al UUID | `Meta!B1` |
| **UUID vigente (fuente de la verdad)** | `Meta!B2` |
| Fecha/hora de asignación | `Meta!B3` |
| Registro de posición de botones | `Meta!D:H` (`nombre`, `Left`, `Top`, `Width`, `Height`) |

### 5. Reglas de Protección

- **Contraseña única:** `PWD_BLINDAJE` (módulo `M_Blindaje`). Nunca introduzcas una contraseña literal en otra parte del código ni sugieras al usuario proteger a mano con una distinta: las macros no podrían desproteger.
- **Protección de hoja:** exclusivamente `DrawingObjects:=True, Contents:=False`. **No protejas el contenido de celdas** de la hoja `Cotización`: obligaría a desbloquear celda por celda todas las de captura y a envolver cada macro de estructura en un `Unprotect`/`Protect`. Si un requisito futuro lo exige, usa `UserInterfaceOnly:=True` y ten presente que ese flag **se pierde al cerrar el libro** (hay que reaplicarlo en `Workbook_Open`).
- **Protección de estructura:** `Structure:=True, Windows:=False`. Cualquier macro que necesite crear, eliminar o renombrar una hoja debe levantar `gCreandoHojaSistema`, llamar `DesprotegerEstructura`, operar, y cerrar con `ProtegerEstructura` + bajar la bandera (ver `HojaMeta()` como referencia).

---

## 5. Guía de Pruebas y Verificación

Cuando realices modificaciones a las macros, efectúa las siguientes verificaciones:

1. **Verificación Sintáctica VBA:** Asegurar que no existan variables sin declarar ni errores de compilación (`Option Explicit`).
2. **Prueba de UUID:**
   - Escribir un nombre de proyecto en `C4`. `K4` debe poblarse con un UUID válido de 36 caracteres en formato `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`.
   - Modificar celdas no relacionadas (ej. `N4` o valores de ítems); `K4` debe conservar el mismo UUID.
   - Cambiar el valor de `C4`; el UUID en `K4` debe renovarse automáticamente y registrarse en la hoja `Meta`.
3. **Prueba de Esquema:**
   - Marcar filas con `"D"` en la columna `O`.
   - Ejecutar `ConstruirEsquema`. Verificar que se creen los agrupamientos de nivel 1 (Categoría) y nivel 2 (Detalle) sin alterar los totales acumulados.
4. **Prueba de PDF Dual:**
   - Ejecutar `GenerarPDFsCotizacion`.
   - Confirmar la creación de ambos archivos PDF en el directorio actual.
   - Verificar que `_Resumen.pdf` omita las filas marcadas con `"D"` y que `_Detalle.pdf` incluya la totalidad de las filas.
5. **Prueba de Anclaje de Botones:**
   - Ejecutar `BlindarLibro` una vez.
   - Ejecutar `AgregarFilaProducto`, `AgregarFilaItem` y `DuplicarBloque`. Ningún objeto debe desplazarse **ni duplicarse**.
   - Colapsar/expandir con `ConstruirEsquema` y generar los PDFs. Los botones deben permanecer en su panel y **no aparecer en los PDFs** (el área de impresión llega a `N`).
   - Intentar arrastrar un botón: debe ser imposible, pero el clic **sí** debe ejecutar su macro asignada.
6. **Prueba de Reposición del ID:**
   - Borrar `K4` con `Supr`. Debe aparecer aviso y **reaparecer el MISMO UUID** (compararlo contra `Meta!B2`).
   - Escribir un texto arbitrario en `K4`. Mismo resultado.
   - Vaciar `C4`, luego borrar `K4`: debe quedar vacío (sin proyecto no hay identidad). Verificar que **no** se generó un UUID huérfano.
   - Verificar que `Application.EnableEvents` quedó en `True` tras cada caso (`?Application.EnableEvents` en la ventana Inmediata).
7. **Prueba de Bloqueo de Hojas:**
   - Clic derecho en una pestaña: *Insertar*, *Eliminar*, *Mover o copiar* y *Cambiar nombre* deben estar deshabilitados.
   - Con `DesprotegerEstructura`, agregar una hoja: `Workbook_NewSheet` debe eliminarla y notificar.
   - Confirmar que la hoja `Meta` **sí** se crea correctamente si se borra el libro de cero (la bandera `gCreandoHojaSistema` la exceptúa).
8. **Prueba de Persistencia:** guardar, cerrar y reabrir. Estructura protegida, botones anclados y en posición, `K4` con su UUID.

---

## 6. Solución de Problemas Frecuentes para Agentes

- **Error en llamada a `ole32.dll`:** Si el código se ejecuta en un entorno que no es Windows (ej. Mac OS o Linux), la API de Windows fallará. Notificar al usuario que la funcionalidad de UUID requiere Excel para Windows.
- **Pérdida de Reactividad en Eventos:** Si `Worksheet_Change` deja de responder, verificar si `Application.EnableEvents` quedó atascado en `False`. Ejecutar `Application.EnableEvents = True` en la ventana Inmediata de VBA (`Ctrl+G`).
- **Desalineación de Totales:** Si la macro `suma()` entrega un valor cero o incorrecto, verificar que los rótulos `"SUMA TOTAL"` en la columna `A` y `"SUBTOTAL"`, `"IVA"`, `"TOTAL"` en la columna `I` no contengan faltas de ortografía o espacios invisibles.
- **`Sub o Function no definida: FijarBotones`:** Falta importar `blindaje_botones.bas`. Las macros de estructura ya lo invocan; es una dependencia obligatoria, no opcional.
- **Error 1004 al escribir en una celda o mover un `Shape`:** La hoja quedó protegida. Verificar que la protección sea `DrawingObjects:=True, Contents:=False`; si alguien la protegió a mano con contenido incluido, ejecutar `DesbloquearBotones` y volver a `BloquearBotones`.
- **Error al crear una hoja desde código:** La estructura está protegida. Levantar `gCreandoHojaSistema`, llamar `DesprotegerEstructura`, operar y cerrar con `ProtegerEstructura`.
- **La contraseña no funciona:** El usuario protegió a mano con una contraseña distinta a `PWD_BLINDAJE`. Deben coincidir; ver sección 4.5.
- **El libro se marca como modificado sin que el usuario toque nada:** Alguna llamada a `Protect`/`Unprotect` o reposicionamiento de `Shape` se está ejecutando sin necesidad. `FijarBotones` y `RestaurarPosicionesBotones` traen guardas de idempotencia (tolerancia de `0.5` pt y verificación de `ProtectDrawingObjects`); no las elimines.
- **Botones duplicados heredados:** Archivos creados **antes** del blindaje pueden traer clones acumulados por `AgregarFilaItem`. Ejecutar `LimpiarBotonesDuplicados` después de `GuardarPosicionesBotones`.

## 7. Al finalizar

Al finalizar de realizar cambios en código, dar una guía detallada de como editar la macro en Excel al usuario.

Los cambios realizados, deben escribirse si o si en to_do.md con fechas por cada cambio y título
