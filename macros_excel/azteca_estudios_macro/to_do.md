# Registro de Tareas y Cambios — Azteca Estudios Macro

## [2026-10-12] Implementación de Mejoras de Robustez, UX y Blindaje Operativo (Plan de Mejora Integral)

Basado en el documento [EVALUACION_Y_PLAN_MEJORA_MACRO.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/EVALUACION_Y_PLAN_MEJORA_MACRO.md) y las reglas operativas de [AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/AGENTS.md).

### 1. Validación de Contexto de Celda Activa (`ValidarFilaEnBloque`)
- **Archivos modificados:**
  - [blindaje_botones.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/blindaje_botones.vb)
  - [bas/blindaje_botones.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/blindaje_botones.bas)
  - [agregar_fila_producto.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_producto.vb)
  - [agregar_fila_fecha.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_fecha.vb)
  - [agregar_fila_item.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/agregar_fila_item.vb)
  - [borrar_fila_seleccionada.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/borrar_fila_seleccionada.vb)
  - [duplicar_bloque.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/duplicar_bloque.vb)
- **Descripción:** Se implementó la función guardiana pública `ValidarFilaEnBloque(fila As Long)` dentro de `M_Blindaje`. Comprueba que la hoja activa sea `Cotización`, que la fila seleccionada sea $\ge \text{FILA\_INICIO}$ y estrictamente anterior al bloque de totales finales (`SUBTOTAL`/`IVA`/`TOTAL`). Si el usuario hace clic en los encabezados o en los totales, se muestra un mensaje informativo amigable y se aborta la operación sin alterar la hoja.

### 2. Blindaje y Manejo Robusto de Excepciones en PDF (`M_PDF`)
- **Archivos modificados:**
  - [generar_pdfs_resumen_detalle.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/generar_pdfs_resumen_detalle.vb)
  - [bas/generar_pdfs_resumen_detalle.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/generar_pdfs_resumen_detalle.bas)
- **Descripción:**
  - **Detección de PDF en uso / bloqueado:** Captura el error de tiempo de ejecución `1004` al exportar. Si el PDF ya está abierto en Acrobat o el navegador, informa amigablemente al usuario para que lo cierre, evitando la pantalla de depuración de VBA.
  - **Validación de libro guardado:** Comprueba `Len(ThisWorkbook.Path) > 0` antes de proceder.
  - **Recálculo automático previo:** Invoca `suma()` antes de exportar para asegurar que los PDFs reflejen los datos más recientes.
  - **Apertura interactiva de carpeta:** Al completar con éxito la exportación de ambos PDFs, ofrece abrir el Explorador de Windows con el archivo seleccionado mediante `Shell explorer.exe /select...`.

### 3. Limpieza Segura con Confirmación y Recálculo (`limpiar_cotizacion.vb`)
- **Archivos modificados:**
  - [limpiar_cotizacion.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/limpiar_cotizacion.vb)
- **Descripción:**
  - Cuadro de confirmación modal `MsgBox vbYesNo` para evitar borrado accidental.
  - Anclaje preventivo de botones (`FijarBotones` y `RestaurarPosicionesBotones`).
  - Protección de la columna A: se conservan intactos los títulos de categoría y la estructura. Solo se limpian las columnas de captura de datos (B a L y O).
  - Recálculo automático con `Call suma` para dejar los totales en $0.00.
  - Corrección de la errata tipográfica en el mensaje de éxito final.

### 4. Recálculo Matemático Tolerante (`suma.vb`)
- **Archivos modificados:**
  - [suma.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/suma.vb)
- **Descripción:** Normalización de cadenas de búsqueda eliminando espacios, puntos y dos puntos (`Replace(Replace(UCase(texto), ":", ""), " ", "") = "SUMATOTAL"`) para tolerar variaciones accidentales de formato en las etiquetas. Distinción explícita de `TOTAL` frente a `SUBTOTAL` y cálculo redondeado de IVA.

### 5. Protección de Folio UUID ante Cambios de Proyecto (`M_UUID`)
- **Archivos modificados:**
  - [asignar_uuid_cotizacion.vb](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/asignar_uuid_cotizacion.vb)
  - [bas/asignar_uuid_cotizacion.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/asignar_uuid_cotizacion.bas)
- **Descripción:** En caso de que se modifique `C4` (Nombre del Proyecto) y ya existiera un UUID asignado en `Meta!B2`, se solicita confirmación al usuario antes de regenerar el folio. Esto permite corregir errores tipográficos en el nombre del proyecto conservando el mismo ID de cotización.

### 6. Sincronización Integral de Módulos `.bas` (cp1252 + CRLF)
- **Archivos modificados:**
  - [bas/asignar_uuid_cotizacion.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/asignar_uuid_cotizacion.bas) (`M_UUID`)
  - [bas/generar_pdfs_resumen_detalle.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/generar_pdfs_resumen_detalle.bas) (`M_PDF`)
  - [bas/blindaje_botones.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/blindaje_botones.bas) (`M_Blindaje`)
  - [bas/esquema_categorias.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/esquema_categorias.bas) (`M_Esquema`)
  - [bas/macro_generadora_uuid.bas](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/bas/macro_generadora_uuid.bas)
- **Descripción:** Regeneración de todos los módulos `.bas` en codificación Windows-1252 con saltos de línea CRLF y atributos `VB_Name` correspondientes para importación directa en el editor VBA de Excel.

### 7. Configuración de Entorno Conda en Documentación de Agentes
- **Archivos modificados:**
  - [AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/AGENTS.md)
  - [.agents/AGENTS.md](file:///e:/Users/1167486/Local/scripts/macros/macros_excel/azteca_estudios_macro/.agents/AGENTS.md)
- **Descripción:** Se documentó la instrucción para activar el entorno de Anaconda `skills_env` cuando se requiera ejecutar herramientas en Python.
